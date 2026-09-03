import "fake-indexeddb/auto";
import { describe, expect, it, vi } from "vitest";
import { AgentContext } from "../src/context";
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
    dbName: `RuntimeToolEventsDB_${nsCounter}`,
    dbVersion: 1,
    localStoragePrefix: `tool-events-test-${nsCounter}`,
    documentSettingsPrefix: `tool-events-test-${nsCounter}`,
    documentIdSettingsKey: `tool-events-test-${nsCounter}-document-id`,
  };
}

function createRuntime(onToolResultSpy?: (id: string, res: string, isError: boolean) => void) {
  const ns = freshNamespace();
  const adapter: RuntimeAdapter = {
    tools: [],
    buildSystemPrompt: () => "Test prompt",
    getDocumentId: async () => "doc-tool-events",
    storageNamespace: ns,
    onToolResult: onToolResultSpy,
  };
  const ctx = new AgentContext({ namespace: ns });
  return new AgentRuntime(adapter, ctx);
}

describe("AgentRuntime Tool Events", () => {
  it("transitions part status to running on tool_execution_start", async () => {
    const runtime = createRuntime();
    await runtime.init();

    const handleEvent = (runtime as unknown as { handleAgentEvent: (e: unknown) => void }).handleAgentEvent;

    // Simulate assistant message containing a tool call
    handleEvent({
      type: "message_start",
      message: {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: "call_abc_1",
            name: "calculate",
            arguments: { expr: "1+1" },
          },
        ],
        timestamp: Date.now(),
      },
    });

    let state = runtime.getState();
    const initialPart = state.messages[0].parts[0];
    expect(initialPart.type).toBe("toolCall");
    if (initialPart.type === "toolCall") {
      expect(initialPart.status).toBe("pending");
    }

    // Fire tool_execution_start
    handleEvent({
      type: "tool_execution_start",
      toolCallId: "call_abc_1",
      toolName: "calculate",
      args: { expr: "1+1" },
    });

    state = runtime.getState();
    const runningPart = state.messages[0].parts[0];
    expect(runningPart.type).toBe("toolCall");
    if (runningPart.type === "toolCall") {
      expect(runningPart.status).toBe("running");
    }

    runtime.dispose();
  });

  it("extracts multimodal image blocks on tool_execution_end", async () => {
    const runtime = createRuntime();
    await runtime.init();

    const handleEvent = (runtime as unknown as { handleAgentEvent: (e: unknown) => void }).handleAgentEvent;

    handleEvent({
      type: "message_start",
      message: {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: "call_img_1",
            name: "screenshot",
            arguments: {},
          },
        ],
        timestamp: Date.now(),
      },
    });

    handleEvent({
      type: "tool_execution_end",
      toolCallId: "call_img_1",
      toolName: "screenshot",
      isError: false,
      result: {
        content: [
          { type: "text", text: "Screenshot captured" },
          { type: "image", data: "iVBORw0KGgoAAAANSUhEUg==", mimeType: "image/png" },
        ],
      },
    });

    const state = runtime.getState();
    const part = state.messages[0].parts[0];
    expect(part.type).toBe("toolCall");
    if (part.type === "toolCall") {
      expect(part.status).toBe("complete");
      expect(part.result).toBe("Screenshot captured");
      expect(part.images).toBeDefined();
      expect(part.images).toHaveLength(1);
      expect(part.images![0]).toEqual({
        data: "iVBORw0KGgoAAAANSUhEUg==",
        mimeType: "image/png",
      });
    }

    runtime.dispose();
  });

  it("suppresses onToolResult when error occurs or followMode is false", async () => {
    const toolResultSpy = vi.fn();
    const runtime = createRuntime(toolResultSpy);
    await runtime.init();

    const handleEvent = (runtime as unknown as { handleAgentEvent: (e: unknown) => void }).handleAgentEvent;

    // FollowMode is enabled by default or via config
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

    // 1. Tool execution with isError: true -> onToolResult must NOT be called
    handleEvent({
      type: "message_start",
      message: {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: "call_err_1",
            name: "delete_sheet",
            arguments: {},
          },
        ],
        timestamp: Date.now(),
      },
    });

    handleEvent({
      type: "tool_execution_end",
      toolCallId: "call_err_1",
      toolName: "delete_sheet",
      isError: true,
      result: "Sheet not found",
    });

    expect(toolResultSpy).not.toHaveBeenCalled();

    // 2. FollowMode is false -> onToolResult must NOT be called even on success
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: false,
      expandToolCalls: false,
    });

    handleEvent({
      type: "message_start",
      message: {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: "call_succ_1",
            name: "read_cell",
            arguments: {},
          },
        ],
        timestamp: Date.now(),
      },
    });

    handleEvent({
      type: "tool_execution_end",
      toolCallId: "call_succ_1",
      toolName: "read_cell",
      isError: false,
      result: "A1 value",
    });

    expect(toolResultSpy).not.toHaveBeenCalled();

    // 3. FollowMode true and no error -> onToolResult MUST be called
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

    handleEvent({
      type: "message_start",
      message: {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: "call_succ_2",
            name: "select_cell",
            arguments: {},
          },
        ],
        timestamp: Date.now(),
      },
    });

    handleEvent({
      type: "tool_execution_end",
      toolCallId: "call_succ_2",
      toolName: "select_cell",
      isError: false,
      result: "Selected B2",
    });

    expect(toolResultSpy).toHaveBeenCalledTimes(1);
    expect(toolResultSpy).toHaveBeenCalledWith("call_succ_2", "Selected B2", false);

    runtime.dispose();
  });
});
