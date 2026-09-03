import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mockOfficeHost } from "../../core/tests/helpers/office-mock";
import { createExecuteOfficeJsTool } from "../src/lib/tools/execute-office-js";

describe("execute_office_js (Word) tool", () => {
  let mockContext: any;
  let agentContext: any;
  let restoreHost: () => void;

  beforeEach(() => {
    agentContext = {
      readFile: vi.fn().mockResolvedValue("sample text"),
      readFileBuffer: vi.fn().mockResolvedValue(new Uint8Array([1, 2, 3])),
      writeFile: vi.fn().mockResolvedValue(undefined),
    };

    mockContext = {
      document: {
        body: {
          text: "Document Body",
          load: vi.fn(),
          clear: vi.fn(),
          delete: vi.fn(),
          paragraphs: {
            getFirst: vi.fn(() => ({
              delete: vi.fn(),
            })),
          },
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    const host = mockOfficeHost("Word", mockContext);
    restoreHost = host.restore;
  });

  afterEach(() => {
    restoreHost();
  });

  it("executes code inside Word.run and returns the result", async () => {
    const tool = createExecuteOfficeJsTool(agentContext);
    const result = await tool.execute("call-word-1", {
      code: "return 10 * 4;",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed.result).toBe(40);
  });

  it("does NOT expose the global Office object inside sandbox (preventing SSO token leakage)", async () => {
    // Ensure Office exists on globalThis in the host
    (globalThis as any).Office = {
      auth: { getAccessToken: vi.fn().mockResolvedValue("secret-token") },
    };

    const tool = createExecuteOfficeJsTool(agentContext);
    const result = await tool.execute("call-word-office-leak", {
      code: "return typeof Office;",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed.result).toBe("undefined");
  });

  it("exposes strictly scoped Word and context objects to code", async () => {
    const tool = createExecuteOfficeJsTool(agentContext);
    const result = await tool.execute("call-word-scope", {
      code: "return { hasContext: !!context, hasWord: typeof Word !== 'undefined' };",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed.result).toEqual({ hasContext: true, hasWord: true });
  });

  it("rejects code containing .delete() by default", async () => {
    const tool = createExecuteOfficeJsTool(agentContext);
    const result = await tool.execute("call-word-delete", {
      code: "context.document.body.paragraphs.getFirst().delete();",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("Destructive operations");
  });

  it("rejects code containing .clear() by default", async () => {
    const tool = createExecuteOfficeJsTool(agentContext);
    const result = await tool.execute("call-word-clear", {
      code: "context.document.body.clear();",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("Destructive operations");
  });

  it("permits destructive operations when allowDestructiveOps is true", async () => {
    const tool = createExecuteOfficeJsTool(agentContext);
    const result = await tool.execute("call-word-destructive-ok", {
      code: "context.document.body.clear(); return 'cleared';",
      allowDestructiveOps: true,
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed.result).toBe("cleared");
  });

  it("permits destructive operations when configured in tool options", async () => {
    const tool = createExecuteOfficeJsTool(agentContext, {
      allowDestructiveOps: true,
    });
    const result = await tool.execute("call-word-opt-ok", {
      code: "return 'allowed';",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed.result).toBe("allowed");
  });

  it("times out if execution hangs beyond the timeout threshold", async () => {
    mockOfficeHost("Word", mockContext, {
      runImpl: () => new Promise((resolve) => setTimeout(resolve, 500)),
    });

    const tool = createExecuteOfficeJsTool(agentContext, { timeoutMs: 50 });
    const result = await tool.execute("call-word-timeout", {
      code: "return 'never';",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("timed out after 0.05s");
  });
});
