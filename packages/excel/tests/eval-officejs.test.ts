import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mockOfficeHost } from "../../core/tests/helpers/office-mock";
import { createEvalOfficeJsTool } from "../src/lib/tools/eval-officejs";

describe("eval_officejs tool", () => {
  let mockContext: any;
  let agentContext: any;
  let restoreHost: () => void;

  beforeEach(() => {
    agentContext = {
      readFile: vi.fn().mockResolvedValue(""),
      readFileBuffer: vi.fn().mockResolvedValue(new Uint8Array()),
      writeFile: vi.fn().mockResolvedValue(undefined),
    };

    const mockSheet = {
      id: "sheet-guid-1",
      name: "Sheet1",
      load: vi.fn(),
      getRange: vi.fn((addr: string) => ({
        address: `Sheet1!${addr}`,
        values: [["Original"]],
        load: vi.fn(),
        clear: vi.fn(),
        delete: vi.fn(),
      })),
    };

    mockContext = {
      workbook: {
        worksheets: {
          items: [mockSheet],
          load: vi.fn(),
          getItem: vi.fn(() => mockSheet),
          getActiveWorksheet: vi.fn(() => mockSheet),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    const host = mockOfficeHost("Excel", mockContext);
    restoreHost = host.restore;
  });

  afterEach(() => {
    restoreHost();
  });

  it("executes code inside Excel.run and returns the result", async () => {
    const tool = createEvalOfficeJsTool(agentContext);
    const result = await tool.execute("call-1", {
      code: "return 1 + 2;",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed.result).toBe(3);
  });

  it("rejects code containing .delete() by default", async () => {
    const tool = createEvalOfficeJsTool(agentContext);
    const result = await tool.execute("call-2", {
      code: "const sheet = context.workbook.worksheets.getActiveWorksheet(); sheet.getRange('A1').delete();",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("Destructive operations");
    expect(parsed.error).toContain("allowDestructiveOps");
  });

  it("rejects code containing .clear() by default", async () => {
    const tool = createEvalOfficeJsTool(agentContext);
    const result = await tool.execute("call-3", {
      code: "context.workbook.worksheets.getActiveWorksheet().getRange('A1').clear();",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("Destructive operations");
  });

  it("rejects code containing worksheet deletion by default", async () => {
    const tool = createEvalOfficeJsTool(agentContext);
    const result = await tool.execute("call-4", {
      code: "const sheet = context.workbook.worksheets.getItem('Sheet1'); // delete worksheet\nsheet.delete();",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("Destructive operations");
  });

  it("allows destructive operations when allowDestructiveOps is true", async () => {
    const tool = createEvalOfficeJsTool(agentContext);
    const result = await tool.execute("call-5", {
      code: "const range = context.workbook.worksheets.getActiveWorksheet().getRange('A1'); range.clear(); return 'cleared';",
      allowDestructiveOps: true,
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed.result).toBe("cleared");
  });

  it("allows destructive operations when configured via tool options", async () => {
    const tool = createEvalOfficeJsTool(agentContext, {
      allowDestructiveOps: true,
    });
    const result = await tool.execute("call-6", {
      code: "return 'allowed by options';",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed.result).toBe("allowed by options");
  });

  it("times out if execution hangs beyond timeout threshold", async () => {
    mockOfficeHost("Excel", mockContext, {
      runImpl: () => new Promise((resolve) => setTimeout(resolve, 500)), // hangs for 500ms
    });

    const tool = createEvalOfficeJsTool(agentContext, { timeoutMs: 50 });
    const result = await tool.execute("call-timeout", {
      code: "return 'never';",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("timed out after 0.05s");
  });

  it("handles primitive, array, and null results without crashing types wrapper", async () => {
    const tool = createEvalOfficeJsTool(agentContext);

    // Number
    let res = await tool.execute("c-num", { code: "return 42;" });
    expect(JSON.parse((res.content[0] as { text: string }).text).result).toBe(
      42,
    );

    // Array
    res = await tool.execute("c-arr", { code: "return [1, 2, 3];" });
    expect(
      JSON.parse((res.content[0] as { text: string }).text).result,
    ).toEqual([1, 2, 3]);

    // Null
    res = await tool.execute("c-null", { code: "return null;" });
    expect(
      JSON.parse((res.content[0] as { text: string }).text).result,
    ).toBeNull();
  });
});
