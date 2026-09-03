import { describe, expect, it } from "vitest";
import { parseSseForId } from "../src/mcp";

describe("MCP SSE parsing", () => {
  it("parses single JSON-RPC message from data lines", () => {
    const sse = `event: message\ndata: {"jsonrpc":"2.0","id":1,"result":{"tools":[]}}\n\n`;
    const parsed = parseSseForId(sse, 1) as { id: number; result: { tools: unknown[] } };
    expect(parsed).toBeDefined();
    expect(parsed.id).toBe(1);
    expect(parsed.result.tools).toEqual([]);
  });

  it("handles multiline data and chunks with keepalives/comments", () => {
    const sse = [
      ": keepalive comment",
      "",
      "event: ping",
      "data:",
      "",
      ": another comment",
      'data: {"jsonrpc":"2.0","id":2,"result":{"content":[{"type":"text","text":"hello"}]}}',
      "",
      ": trailing comment",
    ].join("\n");

    const parsed = parseSseForId(sse, 2) as { id: number; result: { content: Array<{ text: string }> } };
    expect(parsed).toBeDefined();
    expect(parsed.id).toBe(2);
    expect(parsed.result.content[0].text).toBe("hello");
  });

  it("matches specific request IDs when multiple events are present", () => {
    const sse = [
      'data: {"jsonrpc":"2.0","id":1,"result":"first"}',
      'data: {"jsonrpc":"2.0","id":2,"result":"second"}',
      'data: {"jsonrpc":"2.0","id":3,"result":"third"}',
    ].join("\n");

    const match2 = parseSseForId(sse, 2) as { id: number; result: string };
    expect(match2).toBeDefined();
    expect(match2.id).toBe(2);
    expect(match2.result).toBe("second");

    const match1 = parseSseForId(sse, 1) as { id: number; result: string };
    expect(match1).toBeDefined();
    expect(match1.id).toBe(1);
    expect(match1.result).toBe("first");
  });

  it("handles error messages and malformed lines gracefully", () => {
    const sse = [
      "data: not valid json",
      'data: {"jsonrpc":"2.0","id":5,"error":{"code":-32601,"message":"Method not found"}}',
    ].join("\n");

    const parsed = parseSseForId(sse, 5) as { id: number; error: { code: number; message: string } };
    expect(parsed).toBeDefined();
    expect(parsed.id).toBe(5);
    expect(parsed.error.message).toBe("Method not found");
  });

  it("returns last message if target ID is not explicitly found", () => {
    const sse = 'data: {"jsonrpc":"2.0","id":99,"result":"fallback"}\n';
    const parsed = parseSseForId(sse, 100) as { id: number; result: string };
    expect(parsed).toBeDefined();
    expect(parsed.result).toBe("fallback");
  });
});
