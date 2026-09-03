import "fake-indexeddb/auto";
import { describe, expect, it } from "vitest";
import { AgentContext } from "../src/context";
import { AgentRuntime, type RuntimeAdapter } from "../src/runtime";
import { getSession, loadVfsFiles } from "../src/storage/db";

// Polyfill localStorage for Node
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
    dbName: `RuntimeLifecycleDB_${nsCounter}`,
    dbVersion: 1,
    localStoragePrefix: `lifecycle-test-${nsCounter}`,
    documentSettingsPrefix: `lifecycle-test-${nsCounter}`,
    documentIdSettingsKey: `lifecycle-test-${nsCounter}-document-id`,
  };
}

function createRuntime(): AgentRuntime {
  const ns = freshNamespace();
  const adapter: RuntimeAdapter = {
    tools: [],
    buildSystemPrompt: () => "Test prompt",
    getDocumentId: async () => "doc-lifecycle",
    storageNamespace: ns,
  };
  const ctx = new AgentContext({ namespace: ns });
  return new AgentRuntime(adapter, ctx);
}

describe("AgentRuntime Lifecycle Events", () => {
  it("purges streaming message and updates error state on message_end with stopReason error", async () => {
    const runtime = createRuntime();
    await runtime.init();

    const handleEvent = (runtime as unknown as { handleAgentEvent: (e: unknown) => void }).handleAgentEvent;

    // Simulate message_start
    handleEvent({
      type: "message_start",
      message: {
        role: "assistant",
        content: [{ type: "text", text: "Starting..." }],
        timestamp: Date.now(),
      },
    });

    let state = runtime.getState();
    expect(state.messages).toHaveLength(1);
    expect(state.messages[0].role).toBe("assistant");

    // Simulate message_update
    handleEvent({
      type: "message_update",
      message: {
        role: "assistant",
        content: [{ type: "text", text: "Streaming chunks..." }],
        timestamp: Date.now(),
      },
    });

    state = runtime.getState();
    expect(state.messages).toHaveLength(1);

    // Simulate message_end with error stopReason
    handleEvent({
      type: "message_end",
      message: {
        role: "assistant",
        stopReason: "error",
        errorMessage: "Provider quota exceeded",
        timestamp: Date.now(),
      },
    });

    state = runtime.getState();
    // In-flight streaming message must be purged
    expect(state.messages).toHaveLength(0);
    // Error state must be updated
    expect(state.error).toBe("Provider quota exceeded");

    runtime.dispose();
  });

  it("persists session and VFS snapshot to IndexedDB on agent_end", async () => {
    const runtime = createRuntime();
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
    await runtime.init();

    const handleEvent = (
      runtime as unknown as { handleAgentEvent: (e: unknown) => void }
    ).handleAgentEvent;

    // Simulate assistant message start and end
    handleEvent({
      type: "message_start",
      message: {
        role: "assistant",
        content: [{ type: "text", text: "Completed response" }],
        timestamp: Date.now(),
      },
    });

    handleEvent({
      type: "message_end",
      message: {
        role: "assistant",
        stopReason: "stop",
        content: [{ type: "text", text: "Completed response" }],
        timestamp: Date.now(),
      },
    });

    // Write a file to VFS
    const ctx = (runtime as unknown as { context: AgentContext }).context;
    await ctx.writeFile("report.txt", "vfs snapshot content");

    // Trigger agent_end
    handleEvent({ type: "agent_end" });

    // Wait a tick for async onStreamingEnd persistence
    await new Promise((resolve) => setTimeout(resolve, 80));

    const sessionId = runtime.getState().currentSession!.id;
    const ns = (runtime as unknown as { adapter: RuntimeAdapter }).adapter
      .storageNamespace!;
    const persisted = await getSession(ns, sessionId);

    expect(persisted).toBeDefined();
    expect(persisted!.id).toBe(sessionId);

    const savedFiles = await loadVfsFiles(ns, sessionId);
    expect(savedFiles.length).toBeGreaterThan(0);
    const reportFile = savedFiles.find((f) => f.path.endsWith("report.txt"));
    expect(reportFile).toBeDefined();
    expect(new TextDecoder().decode(reportFile!.data)).toBe(
      "vfs snapshot content",
    );

    runtime.dispose();
  });
});
