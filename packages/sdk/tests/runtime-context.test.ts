import "fake-indexeddb/auto";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AgentContext } from "../src/context";
import * as oauthModule from "../src/oauth";
import { saveOAuthCredentials } from "../src/oauth";
import { AgentRuntime, type RuntimeAdapter } from "../src/runtime";

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

let nsCounter = 0;
function freshNamespace() {
  nsCounter++;
  return {
    dbName: `RuntimeContextDB_${nsCounter}`,
    dbVersion: 1,
    localStoragePrefix: `context-test-${nsCounter}`,
    documentSettingsPrefix: `context-test-${nsCounter}`,
    documentIdSettingsKey: `context-test-${nsCounter}-document-id`,
  };
}

describe("AgentRuntime Prompt Context & OAuth Refresh", () => {
  beforeEach(() => {
    localStorage.clear();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("assembles prompt context with <doc_context>, nameMap, and <attachments>", async () => {
    let capturedPrompt = "";
    const ns = freshNamespace();

    const adapter: RuntimeAdapter = {
      tools: [],
      buildSystemPrompt: () => "Test prompt",
      getDocumentId: async () => "doc-123",
      metadataTag: "custom_doc_meta",
      storageNamespace: ns,
      getDocumentMetadata: async () => ({
        metadata: { title: "Quarterly Report", sheets: ["Summary", "Details"] },
        nameMap: { 1: "Sheet1", 2: "Sheet2" },
      }),
    };

    const ctx = new AgentContext({ namespace: ns });
    const runtime = new AgentRuntime(adapter, ctx);
    await runtime.init();

    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    // Mock agent.prompt
    const agent = (runtime as unknown as { agent: { prompt: (text: string) => Promise<void> } }).agent;
    agent.prompt = vi.fn(async (text: string) => {
      capturedPrompt = text;
    });

    await runtime.sendMessage("Please review this file", ["report.pdf", "data.csv"]);

    // Verify <attachments>
    expect(capturedPrompt).toContain("<attachments>\n/home/user/uploads/report.pdf\n/home/user/uploads/data.csv\n</attachments>");

    // Verify metadata tag and JSON content
    expect(capturedPrompt).toContain("<custom_doc_meta>");
    expect(capturedPrompt).toContain('"title": "Quarterly Report"');
    expect(capturedPrompt).toContain("</custom_doc_meta>");

    // Verify original content
    expect(capturedPrompt).toContain("Please review this file");

    // Verify nameMap was updated in state
    expect(runtime.getState().nameMap).toEqual({ 1: "Sheet1", 2: "Sheet2" });
    expect(runtime.getName(1)).toBe("Sheet1");

    runtime.dispose();
  });

  it("refreshes expired OAuth tokens transparently in getActiveApiKey", async () => {
    const ns = freshNamespace();

    const adapter: RuntimeAdapter = {
      tools: [],
      buildSystemPrompt: () => "Test prompt",
      getDocumentId: async () => "doc-oauth",
      storageNamespace: ns,
    };

    const ctx = new AgentContext({ namespace: ns });
    const runtime = new AgentRuntime(adapter, ctx);
    await runtime.init();

    // Save an expired OAuth credential
    saveOAuthCredentials(ns, "google", {
      access: "old-access-token",
      refresh: "valid-refresh-token",
      expires: Date.now() - 60000, // expired 1 minute ago
    });

    runtime.applyConfig({
      provider: "google",
      apiKey: "",
      model: "gemini-2.0-flash",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
      authMethod: "oauth",
    });

    const refreshSpy = vi.spyOn(oauthModule, "refreshOAuthToken").mockResolvedValue({
      access: "new-fresh-access-token",
      refresh: "valid-refresh-token",
      expires: Date.now() + 3600000,
    });

    const getActiveApiKey = (
      runtime as unknown as {
        getActiveApiKey: (cfg: typeof runtime.state.providerConfig) => Promise<string>;
      }
    ).getActiveApiKey.bind(runtime);

    const config = runtime.getState().providerConfig!;
    const activeKey = await getActiveApiKey(config);

    expect(refreshSpy).toHaveBeenCalledWith(
      "google",
      "valid-refresh-token",
      "",
      false,
    );
    expect(activeKey).toBe("new-fresh-access-token");

    // Calling it again should use the cached non-expired token without refreshing
    const activeKey2 = await getActiveApiKey(config);
    expect(refreshSpy).toHaveBeenCalledTimes(1);
    expect(activeKey2).toBe("new-fresh-access-token");

    runtime.dispose();
  });
});
