import "fake-indexeddb/auto";
import { describe, expect, it, vi } from "vitest";
import { AgentContext } from "../src/context";
import { AgentRuntime, type RuntimeAdapter } from "../src/runtime";
import {
  computeArgsHash,
  isToolAllowed,
  type TelemetryEvent,
  type TelemetrySink,
  type ToolAuditEvent,
  type ToolPolicyConfig,
  wrapToolsWithPolicyAndAudit,
} from "../src/telemetry";
import { createBashTool } from "../src/tools/bash";
import { defineTool, toolError, toolSuccess } from "../src/tools/types";
import { Type } from "@sinclair/typebox";

describe("Policy & Audit Telemetry", () => {
  describe("computeArgsHash", () => {
    it("calculates a valid SHA-256 hex string", async () => {
      const hash = await computeArgsHash({ command: "echo hello" });
      expect(hash).toMatch(/^[a-f0-9]{64}$/);
    });

    it("produces known SHA-256 for empty object", async () => {
      const hash = await computeArgsHash({});
      // SHA-256("{}") = 44136fa355b3678a1146ad16f7e8649e94fb4fc21fe77e8310c060f61caaff8a
      expect(hash).toBe(
        "44136fa355b3678a1146ad16f7e8649e94fb4fc21fe77e8310c060f61caaff8a",
      );
    });

    it("produces deterministic hashes regardless of key order", async () => {
      const hashA = await computeArgsHash({ a: "test", b: 123 });
      const hashB = await computeArgsHash({ b: 123, a: "test" });
      expect(hashA).toBe(hashB);
    });

    it("handles undefined and null parameters", async () => {
      const hashNull = await computeArgsHash(null);
      const hashUndef = await computeArgsHash(undefined);
      expect(hashNull).toBe(
        "44136fa355b3678a1146ad16f7e8649e94fb4fc21fe77e8310c060f61caaff8a",
      );
      expect(hashUndef).toBe(hashNull);
    });
  });

  describe("isToolAllowed", () => {
    it("allows all tools when policy is undefined", () => {
      expect(isToolAllowed("bash")).toBe(true);
      expect(isToolAllowed("read_file")).toBe(true);
    });

    it("denies unlisted tools when allowedTools is specified", () => {
      const policy: ToolPolicyConfig = {
        allowedTools: ["read_file", "edit_file"],
      };
      expect(isToolAllowed("read_file", policy)).toBe(true);
      expect(isToolAllowed("edit_file", policy)).toBe(true);
      expect(isToolAllowed("bash", policy)).toBe(false);
      expect(isToolAllowed("eval_officejs", policy)).toBe(false);
    });

    it("denies tools explicitly in deniedTools", () => {
      const policy: ToolPolicyConfig = {
        deniedTools: ["bash", "eval_officejs"],
      };
      expect(isToolAllowed("read_file", policy)).toBe(true);
      expect(isToolAllowed("bash", policy)).toBe(false);
      expect(isToolAllowed("eval_officejs", policy)).toBe(false);
    });

    it("prioritizes deniedTools over allowedTools", () => {
      const policy: ToolPolicyConfig = {
        allowedTools: ["bash", "read_file"],
        deniedTools: ["bash"],
      };
      expect(isToolAllowed("bash", policy)).toBe(false);
      expect(isToolAllowed("read_file", policy)).toBe(true);
    });

    it("denies destructive operations when allowDestructiveOps is false", () => {
      const policy: ToolPolicyConfig = {
        allowDestructiveOps: false,
      };
      expect(isToolAllowed("clear_sheet", policy, true)).toBe(false);
      expect(isToolAllowed("read_sheet", policy, false)).toBe(true);
    });
  });

  describe("wrapToolsWithPolicyAndAudit", () => {
    const mockTool = defineTool({
      name: "echo_tool",
      label: "Echo",
      description: "Echo parameters",
      parameters: Type.Object({ text: Type.String() }),
      execute: async (_id, params) => toolSuccess({ echoed: params.text }),
    });

    const errorTool = defineTool({
      name: "fail_tool",
      label: "Fail",
      description: "Always fails",
      parameters: Type.Object({}),
      execute: async () => {
        throw new Error("Tool explosion");
      },
    });

    const toolErrorTool = defineTool({
      name: "soft_fail_tool",
      label: "Soft Fail",
      description: "Returns toolError",
      parameters: Type.Object({}),
      execute: async () => toolError("Soft error occurred"),
    });

    it("denies unlisted tools, throws an error, and logs denied audit event", async () => {
      const events: TelemetryEvent[] = [];
      const sink: TelemetrySink = {
        emit: (event) => events.push(event),
      };

      const wrapped = wrapToolsWithPolicyAndAudit([mockTool], {
        policy: { allowedTools: ["different_tool"] },
        telemetrySink: sink,
        getDocumentId: () => "doc-123",
        getUserId: () => "user-456",
      });

      expect(wrapped).toHaveLength(1);
      const tool = wrapped[0];

      await expect(
        tool.execute("call_1", { text: "hello" }),
      ).rejects.toThrow("Tool 'echo_tool' is denied by policy");

      expect(events).toHaveLength(1);
      const audit = events[0] as ToolAuditEvent;
      expect(audit.toolCallId).toBe("call_1");
      expect(audit.toolName).toBe("echo_tool");
      expect(audit.status).toBe("denied");
      expect(audit.errorMessage).toContain("denied by policy");
      expect(audit.documentId).toBe("doc-123");
      expect(audit.userId).toBe("user-456");
      expect(audit.argsHash).toMatch(/^[a-f0-9]{64}$/);
      expect(audit.durationMs).toBeGreaterThanOrEqual(0);
      expect(audit.timestamp).toBeGreaterThan(0);
    });

    it("logs audit events on successful execution with duration and argsHash", async () => {
      const events: TelemetryEvent[] = [];
      const sink: TelemetrySink = {
        emit: (event) => events.push(event),
      };

      const wrapped = wrapToolsWithPolicyAndAudit([mockTool], {
        policy: { allowedTools: ["echo_tool"] },
        telemetrySink: sink,
        getDocumentId: () => "doc-abc",
        getUserId: () => "user-xyz",
      });

      const res = await wrapped[0].execute("call_2", { text: "sample" });
      expect(res).toBeDefined();

      expect(events).toHaveLength(1);
      const audit = events[0] as ToolAuditEvent;
      expect(audit.toolCallId).toBe("call_2");
      expect(audit.toolName).toBe("echo_tool");
      expect(audit.status).toBe("success");
      expect(audit.errorMessage).toBeUndefined();
      expect(audit.documentId).toBe("doc-abc");
      expect(audit.userId).toBe("user-xyz");
      expect(audit.argsHash).toBe(
        await computeArgsHash({ text: "sample" }),
      );
      expect(audit.durationMs).toBeGreaterThanOrEqual(0);
    });

    it("logs audit event with status 'error' when execute throws", async () => {
      const events: TelemetryEvent[] = [];
      const sink: TelemetrySink = {
        emit: (event) => events.push(event),
      };

      const wrapped = wrapToolsWithPolicyAndAudit([errorTool], {
        telemetrySink: sink,
      });

      await expect(wrapped[0].execute("call_3", {})).rejects.toThrow(
        "Tool explosion",
      );

      expect(events).toHaveLength(1);
      const audit = events[0] as ToolAuditEvent;
      expect(audit.toolCallId).toBe("call_3");
      expect(audit.toolName).toBe("fail_tool");
      expect(audit.status).toBe("error");
      expect(audit.errorMessage).toBe("Tool explosion");
    });

    it("logs audit event with status 'error' when execute returns toolError", async () => {
      const events: TelemetryEvent[] = [];
      const sink: TelemetrySink = {
        emit: (event) => events.push(event),
      };

      const wrapped = wrapToolsWithPolicyAndAudit([toolErrorTool], {
        telemetrySink: sink,
      });

      const res = await wrapped[0].execute("call_4", {});
      expect(res).toBeDefined();

      expect(events).toHaveLength(1);
      const audit = events[0] as ToolAuditEvent;
      expect(audit.toolCallId).toBe("call_4");
      expect(audit.toolName).toBe("soft_fail_tool");
      expect(audit.status).toBe("error");
      expect(audit.errorMessage).toBe("Soft error occurred");
    });
  });

  describe("AgentRuntime Integration", () => {
    function createMockAdapter(overrides?: Partial<RuntimeAdapter>): RuntimeAdapter {
      return {
        tools: () => [
          defineTool({
            name: "test_tool",
            label: "Test",
            description: "Test tool",
            parameters: Type.Object({ input: Type.String() }),
            execute: async (_id, params) => toolSuccess(params),
          }),
        ],
        buildSystemPrompt: () => "system prompt",
        getDocumentId: vi.fn().mockResolvedValue("doc-runtime-1"),
        getUserId: vi.fn().mockResolvedValue("entra-user-999"),
        storageNamespace: {
          dbName: "TestDB_ws04",
          localStoragePrefix: "test-ws04",
          documentSettingsPrefix: "test-ws04",
        },
        ...overrides,
      };
    }

    it("resolves userId during init() and includes it in emitted audit events", async () => {
      const events: TelemetryEvent[] = [];
      const sink: TelemetrySink = {
        emit: (e) => events.push(e),
      };

      const adapter = createMockAdapter();
      const ctx = new AgentContext({ namespace: adapter.storageNamespace });
      const runtime = new AgentRuntime(adapter, ctx, {
        telemetrySink: sink,
        toolPolicy: { allowedTools: ["test_tool"] },
      });

      await runtime.init();

      expect(runtime.getUserId()).toBe("entra-user-999");
      expect(runtime.getDocumentId()).toBe("doc-runtime-1");

      const tools = runtime.getTools();
      expect(tools).toHaveLength(1);

      await tools[0].execute("call_runtime", { input: "ping" });

      expect(events).toHaveLength(1);
      const audit = events[0] as ToolAuditEvent;
      expect(audit.toolCallId).toBe("call_runtime");
      expect(audit.toolName).toBe("test_tool");
      expect(audit.userId).toBe("entra-user-999");
      expect(audit.documentId).toBe("doc-runtime-1");
      expect(audit.status).toBe("success");
    });

    it("enforces tool policy inside AgentRuntime.getTools()", async () => {
      const events: TelemetryEvent[] = [];
      const sink: TelemetrySink = {
        emit: (e) => events.push(e),
      };

      const adapter = createMockAdapter({
        toolPolicy: { deniedTools: ["test_tool"] },
      });
      const ctx = new AgentContext({ namespace: adapter.storageNamespace });
      const runtime = new AgentRuntime(adapter, ctx, {
        telemetrySink: sink,
      });

      await runtime.init();

      const tools = runtime.getTools();
      await expect(
        tools[0].execute("call_blocked", { input: "bad" }),
      ).rejects.toThrow("Tool 'test_tool' is denied by policy");

      expect(events).toHaveLength(1);
      expect((events[0] as ToolAuditEvent).status).toBe("denied");
    });

    it("policy-gates custom commands inside bash", async () => {
      const ctx = new AgentContext();
      ctx.setCustomCommands(() => ({
        commands: [
          {
            name: "forbidden-cmd",
            execute: async () => ({
              stdout: "unauthorized result",
              stderr: "",
              exitCode: 0,
            }),
          },
        ],
        promptSnippets: ["- forbidden-cmd: test command"],
      }));

      // Create bash tool with policy denying forbidden-cmd
      const bashTool = createBashTool(ctx, {
        policy: { deniedTools: ["forbidden-cmd"] },
      });

      const res = await (bashTool.execute as any)("call_bash", {
        command: "forbidden-cmd",
      });

      const text = res.content[0].text;
      expect(text).toContain("Command 'forbidden-cmd' is denied by policy");
      expect(text).not.toContain("unauthorized result");
    });

    it("updates error state in deleteFile/removeUpload when file deletion fails", async () => {
      const adapter = createMockAdapter();
      const ctx = new AgentContext({ namespace: adapter.storageNamespace });
      const runtime = new AgentRuntime(adapter, ctx);

      // Add a dummy upload to state
      (runtime as any).update({
        uploads: [{ name: "protected.txt", size: 100 }],
      });

      // Mock ctx.deleteFile to throw an error
      vi.spyOn(ctx, "deleteFile").mockRejectedValue(
        new Error("Permission denied deleting file"),
      );

      await runtime.removeUpload("protected.txt");

      // uploads should NOT have been desynchronized (file still present in uploads)
      expect(runtime.getState().uploads).toHaveLength(1);
      expect(runtime.getState().uploads[0].name).toBe("protected.txt");
      // error state should be populated
      expect(runtime.getState().error).toBe(
        "Permission denied deleting file",
      );
    });
  });
});
