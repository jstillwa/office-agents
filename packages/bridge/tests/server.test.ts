import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import net from "node:net";
import https from "node:https";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { WebSocket } from "ws";
import { requestJson } from "../src/http-client";
import {
  createBridgeServer,
  type BridgeServerHandle,
} from "../src/server";
import {
  isBridgeHelloMessage,
  isBridgeInvokeMessage,
  type BridgeResponseMessage,
  type BridgeSessionSnapshot,
  type BridgeWireMessage,
} from "../src/protocol";

const silentLogger = {
  log: () => undefined,
  warn: () => undefined,
  error: () => undefined,
};

function createSnapshot(
  overrides: Partial<BridgeSessionSnapshot> = {},
): BridgeSessionSnapshot {
  const now = Date.now();
  return {
    sessionId: "excel:test-session",
    instanceId: "instance-1",
    app: "excel",
    appName: "Excel",
    appVersion: "1.0.0",
    metadataTag: "doc_context",
    documentId: "doc-123",
    documentMetadata: { sheetCount: 3 },
    tools: [{ name: "echo" }, { name: "eval_officejs" }],
    host: {
      host: "excel",
      platform: "desktop",
      href: "https://localhost/taskpane.html",
      title: "Office Agents",
    },
    connectedAt: now,
    updatedAt: now,
    ...overrides,
  };
}

async function getFreePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      if (!address || typeof address === "string") {
        reject(new Error("Failed to allocate a free port"));
        return;
      }
      const { port } = address;
      server.close((error) => {
        if (error) {
          reject(error);
          return;
        }
        resolve(port);
      });
    });
  });
}

function createTempTlsMaterial() {
  const dir = mkdtempSync(path.join(tmpdir(), "office-bridge-test-"));
  const keyPath = path.join(dir, "localhost.key");
  const certPath = path.join(dir, "localhost.crt");

  execFileSync("openssl", [
    "req",
    "-x509",
    "-newkey",
    "rsa:2048",
    "-keyout",
    keyPath,
    "-out",
    certPath,
    "-nodes",
    "-subj",
    "/CN=localhost",
    "-addext",
    "subjectAltName=DNS:localhost,IP:127.0.0.1",
    "-days",
    "1",
  ]);

  return { dir, keyPath, certPath };
}

async function connectClient(url: string): Promise<WebSocket> {
  return new Promise((resolve, reject) => {
    const socket = new WebSocket(url, {
      rejectUnauthorized: false,
    });
    socket.once("open", () => resolve(socket));
    socket.once("error", reject);
  });
}

async function waitForParsedMessage(socket: WebSocket): Promise<BridgeWireMessage> {
  return new Promise((resolve, reject) => {
    const onMessage = (raw: Buffer) => {
      cleanup();
      resolve(JSON.parse(raw.toString("utf8")) as BridgeWireMessage);
    };
    const onError = (error: Error) => {
      cleanup();
      reject(error);
    };
    const cleanup = () => {
      socket.off("message", onMessage);
      socket.off("error", onError);
    };
    socket.on("message", onMessage);
    socket.on("error", onError);
  });
}

describe("bridge server", () => {
  let tlsDir = "";
  let server: BridgeServerHandle | null = null;
  let socket: WebSocket | null = null;

  beforeEach(() => {
    tlsDir = "";
  });

  afterEach(async () => {
    if (socket) {
      socket.terminate();
      socket = null;
    }
    if (server) {
      await server.close();
      server = null;
    }
    if (tlsDir) {
      rmSync(tlsDir, { recursive: true, force: true });
      tlsDir = "";
    }
  });

  it("registers a session, records bounded event history, and exposes session data over HTTPS", async () => {
    const tls = createTempTlsMaterial();
    tlsDir = tls.dir;
    const port = await getFreePort();
    server = await createBridgeServer({
      host: "127.0.0.1",
      port,
      certPath: tls.certPath,
      keyPath: tls.keyPath,
      eventLimit: 3,
      logger: silentLogger,
    });

    socket = await connectClient(server.wsUrl);
    socket.send(
      JSON.stringify({
        type: "hello",
        role: "office-addin",
        protocolVersion: 1,
        snapshot: createSnapshot(),
      }),
    );

    await waitForParsedMessage(socket);

    socket.send(
      JSON.stringify({ type: "event", event: "selection_changed", ts: 1 }),
    );
    socket.send(JSON.stringify({ type: "event", event: "tool_executed", ts: 2 }));
    socket.send(JSON.stringify({ type: "event", event: "session_updated", ts: 3, payload: createSnapshot({
      documentMetadata: { sheetCount: 4, activeSheet: "Summary" },
      updatedAt: Date.now() + 1,
    }) }));

    const health = (await requestJson(
      "GET",
      "/health",
      undefined,
      { baseUrl: server.httpUrl },
    )) as { ok: boolean; sessions: number };
    const sessionsResponse = (await requestJson(
      "GET",
      "/sessions",
      undefined,
      { baseUrl: server.httpUrl },
    )) as { ok: boolean; sessions: Array<{ snapshot: BridgeSessionSnapshot }> };
    const events = server.getEvents("excel:test-session", 10);

    expect(health.ok).toBe(true);
    expect(health.sessions).toBe(1);
    expect(sessionsResponse.sessions).toHaveLength(1);
    expect(sessionsResponse.sessions[0].snapshot.documentMetadata).toEqual({
      sheetCount: 4,
      activeSheet: "Summary",
    });
    expect(events.map((event) => event.event)).toEqual([
      "selection_changed",
      "tool_executed",
      "session_updated",
    ]);
  });

  it("forwards tool invocations over WebSocket and returns the response to the HTTPS caller", async () => {
    const tls = createTempTlsMaterial();
    tlsDir = tls.dir;
    const port = await getFreePort();
    server = await createBridgeServer({
      host: "127.0.0.1",
      port,
      certPath: tls.certPath,
      keyPath: tls.keyPath,
      logger: silentLogger,
    });

    socket = await connectClient(server.wsUrl);
    socket.send(
      JSON.stringify({
        type: "hello",
        role: "office-addin",
        protocolVersion: 1,
        snapshot: createSnapshot(),
      }),
    );
    await waitForParsedMessage(socket);

    const invokePromise = waitForParsedMessage(socket).then((message) => {
      if (message.type !== "invoke") {
        throw new Error(`Expected invoke message, got ${message.type}`);
      }

      expect(message.method).toBe("execute_tool");
      expect(message.params).toEqual({
        toolName: "echo",
        args: { value: 42, format: "json" },
      });

      const response: BridgeResponseMessage = {
        type: "response",
        requestId: message.requestId,
        ok: true,
        result: {
          toolCallId: "tool_123",
          toolName: "echo",
          isError: false,
          result: {
            content: [{ type: "text", text: "42" }],
          },
          resultText: "42",
          images: [],
        },
      };
      socket?.send(JSON.stringify(response));
    });

    const result = (await requestJson(
      "POST",
      "/sessions/excel%3Atest-session/tools/echo",
      { args: { value: 42, format: "json" } },
      { baseUrl: server.httpUrl },
    )) as {
      ok: boolean;
      result: { resultText: string; toolName: string };
    };

    await invokePromise;

    expect(result.ok).toBe(true);
    expect(result.result.toolName).toBe("echo");
    expect(result.result.resultText).toBe("42");
  });

  it("rejects pending invocations when the WebSocket session disconnects mid-request", async () => {
    const tls = createTempTlsMaterial();
    tlsDir = tls.dir;
    const port = await getFreePort();
    server = await createBridgeServer({
      host: "127.0.0.1",
      port,
      certPath: tls.certPath,
      keyPath: tls.keyPath,
      logger: silentLogger,
    });

    socket = await connectClient(server.wsUrl);
    socket.send(
      JSON.stringify({
        type: "hello",
        role: "office-addin",
        protocolVersion: 1,
        snapshot: createSnapshot(),
      }),
    );
    await waitForParsedMessage(socket);

    const invocation = server.invokeSession({
      sessionId: "excel:test-session",
      method: "ping",
      timeoutMs: 5_000,
    });

    const message = await waitForParsedMessage(socket);
    expect(message.type).toBe("invoke");
    socket.close();

    await expect(invocation).rejects.toThrow(/disconnected/i);
  });

  it("restricts HTTP requests by origin and does not emit wildcard CORS", async () => {
    const tls = createTempTlsMaterial();
    tlsDir = tls.dir;
    const port = await getFreePort();
    server = await createBridgeServer({
      port,
      certPath: tls.certPath,
      keyPath: tls.keyPath,
      logger: silentLogger,
    });

    // Request from allowed origin https://localhost:3000
    const allowedRes = await new Promise<{ statusCode?: number; headers: Record<string, string | string[] | undefined> }>((resolve, reject) => {
      const req = https.request(
        `${server!.httpUrl}/health`,
        {
          method: "GET",
          headers: { Origin: "https://localhost:3000" },
          ca: readFileSync(tls.certPath),
        },
        (res) => {
          res.resume();
          res.on("end", () => resolve({ statusCode: res.statusCode, headers: res.headers }));
        },
      );
      req.on("error", reject);
      req.end();
    });

    expect(allowedRes.statusCode).toBe(200);
    expect(allowedRes.headers["access-control-allow-origin"]).toBe("https://localhost:3000");
    expect(allowedRes.headers["access-control-allow-origin"]).not.toBe("*");

    // Request from disallowed origin https://evil.com is rejected with 403 Forbidden
    const forbiddenRes = await new Promise<{ statusCode?: number; headers: Record<string, string | string[] | undefined> }>((resolve, reject) => {
      const req = https.request(
        `${server!.httpUrl}/health`,
        {
          method: "GET",
          headers: { Origin: "https://evil.com" },
          ca: readFileSync(tls.certPath),
        },
        (res) => {
          res.resume();
          res.on("end", () => resolve({ statusCode: res.statusCode, headers: res.headers }));
        },
      );
      req.on("error", reject);
      req.end();
    });

    expect(forbiddenRes.statusCode).toBe(403);
    expect(forbiddenRes.headers["access-control-allow-origin"]).toBeUndefined();
  });

  it("restricts WebSocket upgrade by origin", async () => {
    const tls = createTempTlsMaterial();
    tlsDir = tls.dir;
    const port = await getFreePort();
    server = await createBridgeServer({
      port,
      certPath: tls.certPath,
      keyPath: tls.keyPath,
      logger: silentLogger,
    });

    // Allowed origin succeeds
    const allowedSocket = new WebSocket(server.wsUrl, {
      headers: { Origin: "https://localhost:3000" },
      ca: readFileSync(tls.certPath),
    });
    await new Promise<void>((resolve, reject) => {
      allowedSocket.once("open", () => resolve());
      allowedSocket.once("error", reject);
    });
    allowedSocket.close();

    // Disallowed origin is rejected
    const evilSocket = new WebSocket(server.wsUrl, {
      headers: { Origin: "https://evil.com" },
      ca: readFileSync(tls.certPath),
    });
    const error = await new Promise<Error>((resolve) => {
      evilSocket.once("error", (err) => resolve(err));
      evilSocket.once("open", () => {
        evilSocket.close();
        resolve(new Error("Expected connection to be rejected"));
      });
    });
    expect(error.message).toMatch(/403|unexpected server response/i);
  });
});

describe("protocol type guards", () => {
  it("isBridgeHelloMessage validates sessionId and documentId strings", () => {
    expect(isBridgeHelloMessage(null)).toBe(false);
    expect(isBridgeHelloMessage(undefined)).toBe(false);
    expect(isBridgeHelloMessage({ type: "hello" })).toBe(false);
    expect(isBridgeHelloMessage({ type: "hello", snapshot: {} })).toBe(false);
    expect(
      isBridgeHelloMessage({
        type: "hello",
        snapshot: { sessionId: 123, documentId: "doc-1" },
      }),
    ).toBe(false);
    expect(
      isBridgeHelloMessage({
        type: "hello",
        snapshot: { sessionId: "sess-1", documentId: 456 },
      }),
    ).toBe(false);
    expect(
      isBridgeHelloMessage({
        type: "hello",
        snapshot: { sessionId: "", documentId: "doc-1" },
      }),
    ).toBe(false);
    expect(
      isBridgeHelloMessage({
        type: "hello",
        snapshot: { sessionId: "sess-1", documentId: "doc-1" },
      }),
    ).toBe(true);
  });

  it("isBridgeInvokeMessage validates requestId and method strings", () => {
    expect(isBridgeInvokeMessage(null)).toBe(false);
    expect(isBridgeInvokeMessage({ type: "invoke" })).toBe(false);
    expect(
      isBridgeInvokeMessage({
        type: "invoke",
        requestId: 123,
        method: "ping",
      }),
    ).toBe(false);
    expect(
      isBridgeInvokeMessage({
        type: "invoke",
        requestId: "req-1",
        method: 456,
      }),
    ).toBe(false);
    expect(
      isBridgeInvokeMessage({
        type: "invoke",
        requestId: "",
        method: "ping",
      }),
    ).toBe(false);
    expect(
      isBridgeInvokeMessage({
        type: "invoke",
        requestId: "req-1",
        method: "ping",
      }),
    ).toBe(true);
  });
});
