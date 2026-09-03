import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { Static, TObject } from "@sinclair/typebox";

export type ToolResult = AgentToolResult<undefined>;

interface ToolConfig<T extends TObject> {
  name: string;
  label: string;
  description: string;
  parameters: T;
  destructive?: boolean;
  execute: (
    toolCallId: string,
    params: Static<T>,
    signal?: AbortSignal,
  ) => Promise<ToolResult>;
}

export function defineTool<T extends TObject>(
  config: ToolConfig<T>,
): AgentTool<T, undefined> {
  return config as AgentTool<T, undefined>;
}

export function toolSuccess(data: unknown): ToolResult {
  const result =
    typeof data === "object" && data !== null ? { ...data } : { result: data };
  return {
    content: [{ type: "text", text: JSON.stringify(result) }],
    details: undefined,
  };
}

export function toolError(message: string): ToolResult {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ success: false, error: message }),
      },
    ],
    details: undefined,
  };
}

export function toolText(text: string): ToolResult {
  return {
    content: [{ type: "text", text }],
    details: undefined,
  };
}
