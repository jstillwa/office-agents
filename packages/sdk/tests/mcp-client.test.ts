import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { McpClient, loadMcpTools, saveMcpConfig } from "../src/mcp";

if (typeof globalThis.localStorage === "undefined") {
  const store: Record<string, string> = {};
  (globalThis as unknown as { localStorage: unknown }).localStorage = {
    getItem: (key: string) => store[key] ?? null,
    setItem: (key: string, value: string) => {
      store[key] = value;
    },
    removeItem: (key: string) => {
      delete store[key];
    },
    clear: () => {
      for (const key of Object.keys(store)) delete store[key];
    },
  };
}

const storageNs = {
  dbName: "test-db",
  sessionKey: "test-session",
  localStoragePrefix: "openexcel",
};

describe("McpClient & loadMcpTools", () => {
  const originalFetch = globalThis.fetch;

  beforeEach(() => {
    localStorage.clear();
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
    vi.restoreAllMocks();
  });

  it("captures mcp-session-id from response and forwards it in subsequent requests", async () => {
    const recordedHeaders: Headers[] = [];
    let callCount = 0;

    globalThis.fetch = vi.fn(async (_url, options) => {
      callCount++;
      const headers = new Headers(options?.headers);
      recordedHeaders.push(headers);

      if (callCount === 1) {
        // initialize response with session id
        return new Response(
          JSON.stringify({
            jsonrpc: "2.0",
            id: 1,
            result: { protocolVersion: "2025-06-18", capabilities: {} },
          }),
          {
            status: 200,
            headers: {
              "content-type": "application/json",
              "mcp-session-id": "session-xyz-123",
            },
          },
        );
      }
      if (callCount === 2) {
        // notification (initialized)
        return new Response(null, { status: 204 });
      }
      // list tools
      return new Response(
        JSON.stringify({
          jsonrpc: "2.0",
          id: 2,
          result: {
            tools: [{ name: "calc", description: "calculator", inputSchema: {} }],
          },
        }),
        {
          status: 200,
          headers: { "content-type": "application/json" },
        },
      );
    }) as unknown as typeof fetch;

    const client = new McpClient("https://mcp.internal.corp/rpc");
    await client.initialize();
    const tools = await client.listTools();

    expect(tools).toHaveLength(1);
    expect(tools[0].name).toBe("calc");

    // Check header forwarding
    expect(recordedHeaders.length).toBe(3);
    // 1st request had no session id
    expect(recordedHeaders[0].get("mcp-session-id")).toBeNull();
    // 2nd request (notification) forwarded session id
    expect(recordedHeaders[1].get("mcp-session-id")).toBe("session-xyz-123");
    // 3rd request forwarded session id
    expect(recordedHeaders[2].get("mcp-session-id")).toBe("session-xyz-123");
  });

  it("sends notifications without id in request body", async () => {
    const recordedBodies: Array<Record<string, unknown>> = [];

    globalThis.fetch = vi.fn(async (_url, options) => {
      recordedBodies.push(JSON.parse(options?.body as string));
      return new Response(
        JSON.stringify({
          jsonrpc: "2.0",
          id: 1,
          result: {},
        }),
        {
          status: 200,
          headers: { "content-type": "application/json" },
        },
      );
    }) as unknown as typeof fetch;

    const client = new McpClient("https://localhost:3000/mcp");
    await client.initialize();

    expect(recordedBodies).toHaveLength(2);
    // initialize had an ID
    expect(recordedBodies[0].id).toBe(1);
    expect(recordedBodies[0].method).toBe("initialize");
    // notifications/initialized had no ID
    expect(recordedBodies[1].id).toBeUndefined();
    expect(recordedBodies[1].method).toBe("notifications/initialized");
  });

  it("wraps tools across text, image, and structured content", async () => {
    globalThis.fetch = vi.fn(async (_url, options) => {
      const body = JSON.parse(options?.body as string);
      if (body.method === "initialize") {
        return new Response(
          JSON.stringify({ jsonrpc: "2.0", id: body.id, result: {} }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (body.method === "notifications/initialized") {
        return new Response(null, { status: 204 });
      }
      if (body.method === "tools/list") {
        return new Response(
          JSON.stringify({
            jsonrpc: "2.0",
            id: body.id,
            result: {
              tools: [
                { name: "text_tool" },
                { name: "image_tool" },
                { name: "structured_tool" },
              ],
            },
          }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      if (body.method === "tools/call") {
        const toolName = body.params?.name;
        if (toolName === "text_tool") {
          return new Response(
            JSON.stringify({
              jsonrpc: "2.0",
              id: body.id,
              result: { content: [{ type: "text", text: "sample result text" }] },
            }),
            { status: 200, headers: { "content-type": "application/json" } },
          );
        }
        if (toolName === "image_tool") {
          return new Response(
            JSON.stringify({
              jsonrpc: "2.0",
              id: body.id,
              result: {
                content: [
                  { type: "image", data: "base64data", mimeType: "image/jpeg" },
                ],
              },
            }),
            { status: 200, headers: { "content-type": "application/json" } },
          );
        }
        if (toolName === "structured_tool") {
          return new Response(
            JSON.stringify({
              jsonrpc: "2.0",
              id: body.id,
              result: { structuredContent: { rows: 5, status: "ok" } },
            }),
            { status: 200, headers: { "content-type": "application/json" } },
          );
        }
      }
      return new Response("Not found", { status: 404 });
    }) as unknown as typeof fetch;

    saveMcpConfig(storageNs, {
      servers: [{ name: "srv", url: "https://localhost:3000/mcp", enabled: true }],
    });

    const tools = await loadMcpTools(storageNs);
    expect(tools).toHaveLength(3);

    const executeTool = async (name: string) => {
      const tool = tools.find((t) => (t as unknown as { name: string }).name === name);
      return (tool?.execute as (id: string, params: Record<string, unknown>) => Promise<unknown>)("call_1", {});
    };

    const textRes = (await executeTool("srv_text_tool")) as { content: Array<{ type: string; text?: string }> };
    expect(textRes.content[0]).toEqual({ type: "text", text: "sample result text" });

    const imgRes = (await executeTool("srv_image_tool")) as { content: Array<{ type: string; data?: string; mimeType?: string }> };
    expect(imgRes.content[0]).toEqual({ type: "image", data: "base64data", mimeType: "image/jpeg" });

    const structRes = (await executeTool("srv_structured_tool")) as { content: Array<{ type: string; text?: string }> };
    expect(structRes.content[0].text).toContain('"rows":5');
  });

  it("recovers gracefully from server errors without throwing in tool execution", async () => {
    globalThis.fetch = vi.fn(async (_url, options) => {
      const body = JSON.parse(options?.body as string);
      if (body.method === "initialize") {
        return new Response(JSON.stringify({ jsonrpc: "2.0", id: body.id, result: {} }), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
      }
      if (body.method === "notifications/initialized") return new Response(null, { status: 204 });
      if (body.method === "tools/list") {
        return new Response(
          JSON.stringify({ jsonrpc: "2.0", id: body.id, result: { tools: [{ name: "failing_tool" }] } }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      // tools/call fails with HTTP 500
      return new Response("Internal Server Error", { status: 500 });
    }) as unknown as typeof fetch;

    saveMcpConfig(storageNs, {
      servers: [{ name: "err_srv", url: "https://localhost:3000/mcp", enabled: true }],
    });

    const tools = await loadMcpTools(storageNs);
    expect(tools).toHaveLength(1);

    const result = (await (tools[0].execute as (id: string, p: Record<string, unknown>) => Promise<unknown>)("call_1", {})) as {
      content: Array<{ type: string; text: string }>;
    };
    expect(result.content[0].type).toBe("text");
    expect(result.content[0].text).toContain("Error: MCP HTTP 500");
  });
});
