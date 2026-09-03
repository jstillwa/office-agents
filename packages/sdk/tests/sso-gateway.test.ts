import "fake-indexeddb/auto";
import { createAssistantMessageEventStream } from "@earendil-works/pi-ai";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AgentContext } from "../src/context";
import { loadMcpConfig, saveMcpConfig } from "../src/mcp";
import { loadOAuthCredentials, saveOAuthCredentials } from "../src/oauth";
import {
  ENTERPRISE_GATEWAY_URL,
  isEnterprise,
  loadSavedConfig,
  type ProviderConfig,
  saveConfig,
} from "../src/provider-config";
import { AgentRuntime, type RuntimeAdapter } from "../src/runtime";
import {
  clearCachedSsoToken,
  parseJwtExpiry,
  resolveOfficeSsoToken,
} from "../src/sso";
import { loadWebConfig, saveWebConfig } from "../src/web/config";

// Mock streamSimple from @earendil-works/pi-ai
const mockStreamSimple = vi.fn();

vi.mock("@earendil-works/pi-ai", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@earendil-works/pi-ai")>();
  return {
    ...actual,
    streamSimple: (...args: any[]) => mockStreamSimple(...args),
  };
});

// Polyfill localStorage for Node
if (typeof globalThis.localStorage === "undefined") {
  const store: Record<string, string> = {};
  (globalThis as any).localStorage = {
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

let testCounter = 0;
function freshNamespace() {
  testCounter++;
  return {
    dbName: `SsoTestDB_${testCounter}`,
    dbVersion: 1,
    localStoragePrefix: `sso-test-${testCounter}`,
    documentSettingsPrefix: `sso-test-${testCounter}`,
    documentIdSettingsKey: `sso-test-${testCounter}-doc-id`,
  };
}

function createMockStream(text = "Hello from corporate gateway") {
  const stream = createAssistantMessageEventStream();
  queueMicrotask(() => {
    const msg = {
      role: "assistant" as const,
      content: [{ type: "text" as const, text }],
      api: "openai-completions",
      provider: "custom",
      model: "corporate-default",
      usage: {
        input: 10,
        output: 5,
        cacheRead: 0,
        cacheWrite: 0,
        totalTokens: 15,
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
      },
      stopReason: "stop" as const,
      timestamp: Date.now(),
    };
    stream.push({ type: "start", partial: { ...msg, content: [] } });
    stream.push({
      type: "text_start",
      contentIndex: 0,
      partial: { ...msg, content: [{ type: "text", text: "" }] },
    });
    stream.push({
      type: "text_delta",
      contentIndex: 0,
      delta: text,
      partial: msg,
    });
    stream.push({
      type: "text_end",
      contentIndex: 0,
      content: text,
      partial: msg,
    });
    stream.push({ type: "done", reason: "stop", message: msg });
    stream.end(msg);
  });
  return stream;
}

function createJwt(expSecondsFromNow: number): string {
  const header = btoa(JSON.stringify({ alg: "RS256", typ: "JWT" }));
  const payload = btoa(
    JSON.stringify({
      aud: "corporate-llm-gateway",
      iss: "https://login.microsoftonline.com/tenant-id/v2.0",
      exp: Math.floor(Date.now() / 1000) + expSecondsFromNow,
      name: "Enterprise User",
      preferred_username: "user@enterprise.org",
    }),
  );
  return `${header}.${payload}.mock-signature`;
}

describe("Enterprise SSO, Corporate LLM Gateway & Storage Lockdown", () => {
  const originalAppMode = process.env.VITE_APP_MODE;

  beforeEach(() => {
    localStorage.clear();
    clearCachedSsoToken();
    vi.clearAllMocks();
  });

  afterEach(() => {
    process.env.VITE_APP_MODE = originalAppMode;
    delete (globalThis as any).OfficeRuntime;
  });

  describe("Zero-BYOK Storage Lockdown in Enterprise Mode", () => {
    beforeEach(() => {
      process.env.VITE_APP_MODE = "enterprise";
    });

    it("verifies isEnterprise returns true", () => {
      expect(isEnterprise()).toBe(true);
    });

    it("verifies zero localStorage writes for apiKey in saveConfig", () => {
      const ns = freshNamespace();
      const config: ProviderConfig = {
        provider: "custom",
        apiKey: "sk-secret-corporate-key-should-never-be-saved",
        model: "gpt-4o",
        useProxy: true,
        proxyUrl: "https://proxy.example.com",
        thinking: "none",
        followMode: true,
        expandToolCalls: false,
        apiType: "openai-completions",
        customBaseUrl: "https://custom.gateway.internal/v1",
      };

      saveConfig(ns, config);

      const raw = localStorage.getItem(`${ns.localStoragePrefix}-provider-config`);
      expect(raw).not.toBeNull();
      const stored = JSON.parse(raw!);
      expect(stored.apiKey).toBeUndefined();
      expect(raw).not.toContain("sk-secret-corporate-key-should-never-be-saved");
      expect(stored.useProxy).toBe(false);
      expect(stored.authMethod).toBe("sso");
    });

    it("loadSavedConfig returns empty apiKey and routes to enterprise gateway in enterprise mode", () => {
      const ns = freshNamespace();
      // Write a legacy config containing an apiKey
      localStorage.setItem(
        `${ns.localStoragePrefix}-provider-config`,
        JSON.stringify({
          provider: "openai",
          apiKey: "legacy-leaked-key",
          model: "gpt-3.5-turbo",
          useProxy: true,
        }),
      );

      const loaded = loadSavedConfig(ns);
      expect(loaded).not.toBeNull();
      expect(loaded?.apiKey).toBe("");
      expect(loaded?.useProxy).toBe(false);
      expect(loaded?.authMethod).toBe("sso");
      expect(loaded?.provider).toBe("custom");
      expect(loaded?.customBaseUrl).toBe(ENTERPRISE_GATEWAY_URL);
    });

    it("verifies zero localStorage writes for oauth-credentials and purges legacy tokens", () => {
      const ns = freshNamespace();
      const oauthKey = `${ns.localStoragePrefix}-oauth-credentials`;
      localStorage.setItem(
        oauthKey,
        JSON.stringify({ anthropic: { access: "stale-token" } }),
      );

      // Loading credentials should purge and return null in enterprise mode
      const loaded = loadOAuthCredentials(ns, "anthropic");
      expect(loaded).toBeNull();
      expect(localStorage.getItem(oauthKey)).toBeNull();

      // Saving credentials should not write to localStorage
      saveOAuthCredentials(ns, "anthropic", {
        access: "new-token",
        refresh: "refresh-token",
        expires: Date.now() + 3600000,
      });
      expect(localStorage.getItem(oauthKey)).toBeNull();
    });

    it("saveMcpConfig eliminates plaintext header persistence in enterprise mode", () => {
      const ns = freshNamespace();
      saveMcpConfig(ns, {
        servers: [
          {
            name: "jira-mcp",
            url: "https://jira.internal/mcp",
            headers: { Authorization: "Bearer secret-token-123" },
          },
        ],
      });

      const raw = localStorage.getItem(`${ns.localStoragePrefix}-mcp-config`);
      expect(raw).not.toBeNull();
      expect(raw).not.toContain("secret-token-123");
      const loaded = loadMcpConfig(ns);
      expect(loaded.servers[0].headers).toBeUndefined();
    });

    it("saveWebConfig strips search/fetch apiKeys in enterprise mode", () => {
      const ns = freshNamespace();
      saveWebConfig(ns, {
        searchProvider: "brave",
        apiKeys: {
          brave: "brave-secret-key",
          serper: "serper-secret-key",
        },
      });

      const loaded = loadWebConfig(ns);
      expect(loaded.apiKeys).toEqual({});
      const raw = localStorage.getItem(`${ns.localStoragePrefix}-web-config`);
      expect(raw).not.toContain("brave-secret-key");
      expect(raw).not.toContain("serper-secret-key");
    });
  });

  describe("Office SSO Entra ID Token Acquisition & Cache", () => {
    it("parses JWT expiry correctly", () => {
      const jwt = createJwt(300); // 300 seconds from now
      const expMs = parseJwtExpiry(jwt);
      expect(expMs).not.toBeNull();
      expect(expMs!).toBeGreaterThan(Date.now() + 290 * 1000);
      expect(expMs!).toBeLessThanOrEqual(Date.now() + 305 * 1000);
    });

    it("acquires Entra bearer token via OfficeRuntime.auth.getAccessToken and caches it", async () => {
      const mockGetAccessToken = vi.fn().mockResolvedValue(createJwt(3600));
      (globalThis as any).OfficeRuntime = {
        auth: { getAccessToken: mockGetAccessToken },
      };

      const token1 = await resolveOfficeSsoToken();
      expect(token1).toContain("mock-signature");
      expect(mockGetAccessToken).toHaveBeenCalledTimes(1);
      expect(mockGetAccessToken).toHaveBeenCalledWith({ allowSignInPrompt: true });

      // Second invocation should use the cached token without re-calling OfficeRuntime
      const token2 = await resolveOfficeSsoToken();
      expect(token2).toBe(token1);
      expect(mockGetAccessToken).toHaveBeenCalledTimes(1);

      // Force refresh should bypass cache and fetch a new token
      mockGetAccessToken.mockResolvedValueOnce(createJwt(1800));
      const token3 = await resolveOfficeSsoToken(true);
      expect(mockGetAccessToken).toHaveBeenCalledTimes(2);
      expect(token3).not.toBe("");
    });
  });

  describe("Runtime SSO Gateway Stream Headers & Suppress x-api-key", () => {
    it("mocks OfficeRuntime.auth.getAccessToken, acquires token, injects Authorization: Bearer <token>, and suppresses x-api-key headers", async () => {
      process.env.VITE_APP_MODE = "enterprise";
      const token = createJwt(3600);
      const mockGetAccessToken = vi.fn().mockResolvedValue(token);
      (globalThis as any).OfficeRuntime = {
        auth: { getAccessToken: mockGetAccessToken },
      };

      mockStreamSimple.mockImplementation(() => createMockStream("Enterprise response"));

      const ns = freshNamespace();
      const adapter: RuntimeAdapter = {
        tools: [],
        buildSystemPrompt: () => "You are an enterprise AI assistant.",
        getDocumentId: async () => "enterprise-doc-42",
        getUserId: async () => "john.doe@enterprise.org",
        storageNamespace: ns,
      };

      const ctx = new AgentContext({ namespace: ns });
      const runtime = new AgentRuntime(adapter, ctx);
      await runtime.init();

      // Ensure config is applied with authMethod = sso
      const config: ProviderConfig = {
        provider: "custom",
        apiKey: "",
        model: "corporate-gpt-4o",
        useProxy: false,
        proxyUrl: "",
        thinking: "none",
        followMode: true,
        expandToolCalls: false,
        apiType: "openai-completions",
        customBaseUrl: "https://gateway.enterprise.corp/v1",
        authMethod: "sso",
      };
      runtime.setProviderConfig(config);

      // Send a message to trigger streaming
      await runtime.sendMessage("Hello Enterprise Gateway");

      // Verify OfficeRuntime was called to acquire Entra ID bearer token
      expect(mockGetAccessToken).toHaveBeenCalledWith({ allowSignInPrompt: true });

      // Verify streamSimple was called with Authorization: Bearer <token>
      expect(mockStreamSimple).toHaveBeenCalled();
      const streamCallArgs = mockStreamSimple.mock.calls[0];
      const streamOptions = streamCallArgs[2];

      // Must have Authorization: Bearer <token>
      expect(streamOptions.headers).toBeDefined();
      expect(streamOptions.headers.Authorization).toBe(`Bearer ${token}`);

      // Must omit options.apiKey so pi-ai does not emit x-api-key headers
      expect(streamOptions.apiKey).toBeUndefined();
      expect(streamOptions.headers["x-api-key"]).toBeUndefined();

      runtime.dispose();
    });

    it("transparently refreshes Entra ID token and retries upon 401 response", async () => {
      process.env.VITE_APP_MODE = "enterprise";
      const token1 = createJwt(3600);
      const token2 = createJwt(7200);
      const mockGetAccessToken = vi
        .fn()
        .mockResolvedValueOnce(token1)
        .mockResolvedValueOnce(token2);

      (globalThis as any).OfficeRuntime = {
        auth: { getAccessToken: mockGetAccessToken },
      };

      let attempt = 0;
      mockStreamSimple.mockImplementation(() => {
        attempt++;
        if (attempt === 1) {
          const error: any = new Error("Unauthorized: 401 token expired");
          error.status = 401;
          throw error;
        }
        return createMockStream("Success on retry with fresh token");
      });

      const ns = freshNamespace();
      const adapter: RuntimeAdapter = {
        tools: [],
        buildSystemPrompt: () => "Enterprise prompt",
        getDocumentId: async () => "doc-401-test",
        storageNamespace: ns,
      };

      const ctx = new AgentContext({ namespace: ns });
      const runtime = new AgentRuntime(adapter, ctx);
      await runtime.init();

      runtime.setProviderConfig({
        provider: "custom",
        apiKey: "",
        model: "corporate-gpt",
        useProxy: false,
        proxyUrl: "",
        thinking: "none",
        followMode: true,
        expandToolCalls: false,
        apiType: "openai-completions",
        customBaseUrl: "https://gateway.enterprise.corp/v1",
        authMethod: "sso",
      });

      await runtime.sendMessage("Trigger token refresh on 401");

      expect(mockGetAccessToken).toHaveBeenCalledTimes(2);
      expect(mockStreamSimple).toHaveBeenCalledTimes(2);

      const secondCallArgs = mockStreamSimple.mock.calls[1];
      expect(secondCallArgs[2].headers.Authorization).toBe(`Bearer ${token2}`);
      expect(secondCallArgs[2].apiKey).toBeUndefined();

      runtime.dispose();
    });
  });

  describe("Storage failure handling and telemetrySink notification", () => {
    it("notifies telemetrySink and updates state.error when storage fails in clearMessages", async () => {
      const ns = freshNamespace();
      const telemetryEvents: any[] = [];
      const adapter: RuntimeAdapter = {
        tools: [],
        buildSystemPrompt: () => "Test prompt",
        getDocumentId: async () => "doc-storage-err",
        telemetrySink: {
          emit: (event) => {
            telemetryEvents.push(event);
          },
        },
        storageNamespace: ns,
      };

      const ctx = new AgentContext({ namespace: ns });
      const runtime = new AgentRuntime(adapter, ctx);
      await runtime.init();

      // Close underlying DB or force failure by invalidating session
      (runtime as any).currentSessionId = "non-existent-session-id";

      // Mock saveSession to reject
      const dbModule = await import("../src/storage/db");
      const spy = vi.spyOn(dbModule, "saveSession").mockRejectedValueOnce(new Error("IDB Disk Full"));

      runtime.clearMessages();

      // Wait for catch handler microtask
      await new Promise((resolve) => setTimeout(resolve, 50));

      expect(runtime.getState().error).toContain("Storage failure: IDB Disk Full");
      expect(telemetryEvents.some((e) => e.type === "storage_error")).toBe(true);

      spy.mockRestore();
      runtime.dispose();
    });
  });
});
