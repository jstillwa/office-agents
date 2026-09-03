import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { CustomCommand } from "just-bash/browser";

export interface ToolPolicyConfig {
  allowedTools?: string[];
  deniedTools?: string[];
  allowDestructiveOps?: boolean;
}

export type ToolAuditStatus = "success" | "error" | "denied";

export interface ToolAuditEvent {
  type?: "tool_audit";
  toolCallId: string;
  toolName: string;
  argsHash: string;
  documentId: string | null;
  userId: string | null;
  timestamp: number;
  durationMs: number;
  status: ToolAuditStatus;
  errorMessage?: string | null;
}

export type TelemetryEvent =
  | ToolAuditEvent
  | ({ type: "tool_audit" } & ToolAuditEvent)
  | { type: string; [key: string]: unknown };

export interface TelemetrySink {
  emit: (event: TelemetryEvent) => void | Promise<void>;
}

export function emitTelemetry(
  sink: TelemetrySink | ((event: TelemetryEvent) => void) | undefined | null,
  event: TelemetryEvent,
): void {
  if (!sink) return;
  try {
    if (typeof sink === "function") {
      sink(event);
    } else if (typeof sink.emit === "function") {
      sink.emit(event);
    }
  } catch (err) {
    console.error("[Telemetry] Error emitting to sink:", err);
  }
}

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") {
    return JSON.stringify(value) ?? "";
  }
  if (Array.isArray(value)) {
    return `[${value.map(stableStringify).join(",")}]`;
  }
  const keys = Object.keys(value as Record<string, unknown>).sort();
  const entries = keys.map(
    (k) =>
      `${JSON.stringify(k)}:${stableStringify((value as Record<string, unknown>)[k])}`,
  );
  return `{${entries.join(",")}}`;
}

export async function computeArgsHash(args: unknown): Promise<string> {
  const json = stableStringify(args ?? {});
  const encoder = new TextEncoder();
  const data = encoder.encode(json);
  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

export function isToolAllowed(
  toolName: string,
  policy?: ToolPolicyConfig,
  isDestructive = false,
): boolean {
  if (!policy) return true;

  if (policy.allowDestructiveOps === false && isDestructive) {
    return false;
  }

  if (policy.deniedTools?.includes(toolName)) {
    return false;
  }

  if (policy.allowedTools && !policy.allowedTools.includes(toolName)) {
    return false;
  }

  return true;
}

export interface WrapToolsOptions {
  policy?: ToolPolicyConfig;
  telemetrySink?: TelemetrySink | ((event: TelemetryEvent) => void);
  getDocumentId?: () => string | null | Promise<string | null>;
  getUserId?: () => string | null | Promise<string | null>;
}

export function wrapToolsWithPolicyAndAudit(
  tools: AgentTool[],
  options: WrapToolsOptions,
): AgentTool[] {
  return tools.map((tool) => {
    const isDestructive =
      (tool as { destructive?: boolean }).destructive === true;

    return {
      ...tool,
      execute: async (toolCallId, params, signal, onUpdate) => {
        const start = performance.now();
        const timestamp = Date.now();
        const argsHash = await computeArgsHash(params);
        const documentId =
          typeof options.getDocumentId === "function"
            ? await options.getDocumentId()
            : null;
        const userId =
          typeof options.getUserId === "function"
            ? await options.getUserId()
            : null;

        const allowed = isToolAllowed(tool.name, options.policy, isDestructive);

        if (!allowed) {
          const durationMs = performance.now() - start;
          const errorMessage = `Tool '${tool.name}' is denied by policy`;
          const auditEvent: ToolAuditEvent = {
            toolCallId,
            toolName: tool.name,
            argsHash,
            documentId: documentId ?? null,
            userId: userId ?? null,
            timestamp,
            durationMs,
            status: "denied",
            errorMessage,
          };
          emitTelemetry(options.telemetrySink, auditEvent);
          throw new Error(errorMessage);
        }

        try {
          const result = await tool.execute(
            toolCallId,
            params,
            signal,
            onUpdate,
          );
          const durationMs = performance.now() - start;

          let isError = false;
          let errorMessage: string | undefined;
          if (result && Array.isArray(result.content)) {
            for (const item of result.content) {
              if (item.type === "text" && typeof item.text === "string") {
                try {
                  const parsed = JSON.parse(item.text);
                  if (parsed && parsed.success === false && parsed.error) {
                    isError = true;
                    errorMessage = parsed.error;
                    break;
                  }
                } catch {
                  // Not JSON
                }
              }
            }
          }

          const auditEvent: ToolAuditEvent = {
            toolCallId,
            toolName: tool.name,
            argsHash,
            documentId: documentId ?? null,
            userId: userId ?? null,
            timestamp,
            durationMs,
            status: isError ? "error" : "success",
            errorMessage: isError ? errorMessage : undefined,
          };
          emitTelemetry(options.telemetrySink, auditEvent);
          return result;
        } catch (error) {
          const durationMs = performance.now() - start;
          const errorMessage =
            error instanceof Error ? error.message : String(error);
          const auditEvent: ToolAuditEvent = {
            toolCallId,
            toolName: tool.name,
            argsHash,
            documentId: documentId ?? null,
            userId: userId ?? null,
            timestamp,
            durationMs,
            status: "error",
            errorMessage,
          };
          emitTelemetry(options.telemetrySink, auditEvent);
          throw error;
        }
      },
    };
  });
}

export function wrapCustomCommandWithPolicy(
  cmd: CustomCommand,
  policy?: ToolPolicyConfig,
): CustomCommand {
  if (!policy) return cmd;

  if ("load" in cmd && typeof cmd.load === "function") {
    return {
      name: cmd.name,
      load: async () => {
        const loaded = await cmd.load();
        return wrapCommandObject(loaded, policy);
      },
    };
  }

  return wrapCommandObject(cmd, policy);
}

function wrapCommandObject(cmd: any, policy?: ToolPolicyConfig): any {
  if (!policy) return cmd;
  return {
    ...cmd,
    execute: async (args: string[], ctx: any) => {
      if (!isToolAllowed(cmd.name, policy)) {
        return {
          stdout: "",
          stderr: `Command '${cmd.name}' is denied by policy\n`,
          exitCode: 126,
        };
      }
      return cmd.execute(args, ctx);
    },
  };
}
