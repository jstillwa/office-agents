import { describe, expect, it } from "vitest";
import { AgentContext } from "../src/context";
import { createEditFileTool } from "../src/tools/edit-file";
import type { ToolResult } from "../src/tools/types";

function parseResult(result: ToolResult): Record<string, unknown> {
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  return JSON.parse(text);
}

describe("edit_file tool", () => {
  it("preserves literal dollar signs without regex token expansion", async () => {
    const ctx = new AgentContext();
    const tool = createEditFileTool(ctx);
    const execute = tool.execute as (
      id: string,
      params: Record<string, unknown>,
    ) => Promise<ToolResult>;

    // Initial file creation
    await execute("call_1", {
      path: "sheet.csv",
      content: "formula,total\nOLD_FORMULA,100\n",
    });

    // Replace with formula containing dollar signs
    const editRes = await execute("call_2", {
      path: "sheet.csv",
      edits: [
        {
          old_text: "OLD_FORMULA",
          new_text: "=$A$1+$B$1",
        },
      ],
    });

    const parsed = parseResult(editRes);
    expect(parsed.success).toBe(true);

    const savedContent = await ctx.readFile("sheet.csv");
    expect(savedContent).toBe("formula,total\n=$A$1+$B$1,100\n");
  });

  it("preserves special replacement patterns like $&, $', $`, and $1", async () => {
    const ctx = new AgentContext();
    const tool = createEditFileTool(ctx);
    const execute = tool.execute as (
      id: string,
      params: Record<string, unknown>,
    ) => Promise<ToolResult>;

    await execute("call_1", {
      path: "test.txt",
      content: "placeholder",
    });

    const editRes = await execute("call_2", {
      path: "test.txt",
      edits: [
        {
          old_text: "placeholder",
          new_text: "$& and $' and $` and $1 and $$",
        },
      ],
    });

    const parsed = parseResult(editRes);
    expect(parsed.success).toBe(true);

    const savedContent = await ctx.readFile("test.txt");
    expect(savedContent).toBe("$& and $' and $` and $1 and $$");
  });

  it("creates and overwrites files with content parameter", async () => {
    const ctx = new AgentContext();
    const tool = createEditFileTool(ctx);
    const execute = tool.execute as (
      id: string,
      params: Record<string, unknown>,
    ) => Promise<ToolResult>;

    const createRes = await execute("call_1", {
      path: "note.txt",
      content: "hello world",
    });
    expect(parseResult(createRes)).toMatchObject({
      success: true,
      action: "created",
    });

    const overwriteRes = await execute("call_2", {
      path: "note.txt",
      content: "new content",
    });
    expect(parseResult(overwriteRes)).toMatchObject({
      success: true,
      action: "overwrote",
    });

    expect(await ctx.readFile("note.txt")).toBe("new content");
  });

  it("fails when path traverses outside /home/user/", async () => {
    const ctx = new AgentContext();
    const tool = createEditFileTool(ctx);
    const execute = tool.execute as (
      id: string,
      params: Record<string, unknown>,
    ) => Promise<ToolResult>;

    const result = await execute("call_1", {
      path: "/skills/prompt.md",
      content: "malicious edit",
    });

    const parsed = parseResult(result);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("outside /home/user/");
  });
});
