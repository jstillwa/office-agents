import type { AgentContext } from "@office-agents/core";
import { sandboxedEval } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import { defineTool, toolError, toolSuccess } from "./types";

/* global Word */

// Destructive-op guardrail against accidental LLM behaviour, NOT a security control
// (trivially bypassed by string concatenation); the real controls are
// the ws-04 tool allow/deny policy and audit log.
const DESTRUCTIVE_PATTERNS = [
  /\.delete\s*\(/,
  /\.clear\s*\(/,
  /\bdelete\b.*worksheet/i,
];

function containsDestructiveOps(code: string): boolean {
  return DESTRUCTIVE_PATTERNS.some((p) => p.test(code));
}

const TIMEOUT_MS = 30_000;

function withTimeout<T>(promise: Promise<T>, ms = TIMEOUT_MS): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return Promise.race([
    promise,
    new Promise<never>((_, reject) => {
      timer = setTimeout(() => {
        reject(new Error(`Execution timed out after ${ms / 1000}s`));
      }, ms);
    }),
  ]).finally(() => {
    if (timer) clearTimeout(timer);
  });
}

export interface ExecuteOfficeJsOptions {
  allowDestructiveOps?: boolean;
  timeoutMs?: number;
}

export function createExecuteOfficeJsTool(
  ctx: AgentContext,
  options?: ExecuteOfficeJsOptions,
) {
  return defineTool({
    name: "execute_office_js",
    label: "Execute Office.js Code",
    description:
      "Execute Office.js JavaScript code to interact with the Word document. " +
      "The code receives a context parameter and runs inside Word.run(). " +
      "Use this for any document operations: inserting/editing text, formatting, tables, images, " +
      "comments, tracked changes, search/replace, OOXML, headers/footers, content controls, and more.",
    parameters: Type.Object({
      code: Type.String({
        description:
          "Async function body that receives 'context: Word.RequestContext'. " +
          "Must call context.sync() to execute batched operations and load() to read properties. " +
          "Return JSON-serializable results. " +
          "readFile(path) returns Promise<string> and readFileBuffer(path) returns Promise<Uint8Array> " +
          "to read files from the virtual filesystem. " +
          "writeFile(path, content) returns Promise<void> to write string or Uint8Array to the virtual filesystem. " +
          "btoa(string) and atob(base64) are available for base64 encoding/decoding.",
      }),
      explanation: Type.Optional(
        Type.String({
          description:
            "Brief explanation of what this code does (max 100 chars)",
          maxLength: 100,
        }),
      ),
      allowDestructiveOps: Type.Optional(
        Type.Boolean({
          description:
            "Set to true to allow destructive operations (.delete(), .clear()). Defaults to false.",
        }),
      ),
    }),
    execute: async (_toolCallId, params) => {
      const allowDestructive =
        params.allowDestructiveOps ?? options?.allowDestructiveOps ?? false;

      if (!allowDestructive && containsDestructiveOps(params.code)) {
        return toolError(
          "Destructive operations (.delete(), .clear()) are not allowed unless allowDestructiveOps is set to true.",
        );
      }

      try {
        const timeoutMs = options?.timeoutMs ?? TIMEOUT_MS;
        const result = await withTimeout(
          Word.run(async (context) => {
            return sandboxedEval(params.code, {
              context,
              Word,
              readFile: (path: string) => ctx.readFile(path),
              readFileBuffer: (path: string) => ctx.readFileBuffer(path),
              writeFile: (path: string, content: string | Uint8Array) =>
                ctx.writeFile(path, content),
            });
          }),
          timeoutMs,
        );

        return toolSuccess({ success: true, result: result ?? null });
      } catch (error) {
        if (error instanceof OfficeExtension.Error) {
          const parts = [error.message];
          if (error.code) parts.push(`Code: ${error.code}`);
          if (error.debugInfo) {
            const { errorLocation, statement, surroundingStatements } =
              error.debugInfo;
            if (errorLocation) parts.push(`Location: ${errorLocation}`);
            if (statement) parts.push(`Statement: ${statement}`);
            if (surroundingStatements?.length)
              parts.push(`Context: ${surroundingStatements.join("; ")}`);
          }
          return toolError(parts.join("\n"));
        }
        const message =
          error instanceof Error
            ? error.message
            : "Unknown error executing code";
        return toolError(message);
      }
    },
  });
}
