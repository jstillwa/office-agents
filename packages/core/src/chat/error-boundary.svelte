<script lang="ts">
  import { emitTelemetry, type TelemetrySink } from "@office-agents/sdk";
  import type { Snippet } from "svelte";

  interface Props {
    children?: Snippet;
    telemetrySink?: TelemetrySink;
  }

  let { children, telemetrySink }: Props = $props();
  let errorMessage = $state("");
  let errorStack = $state<string | undefined>(undefined);
  let copied = $state(false);

  function getOfficeDiagnostics() {
    if (typeof Office !== "undefined" && Office?.context?.diagnostics) {
      const diag = Office.context.diagnostics;
      return {
        host: diag.host,
        platform: diag.platform,
        version: diag.version,
      };
    }
    return null;
  }

  function onerror(error: unknown) {
    const err = error instanceof Error ? error : new Error(String(error));
    errorMessage = err.message || "Something went wrong";
    errorStack = err.stack;
    console.error("[UI] Unhandled render error:", error);

    const officeDiag = getOfficeDiagnostics();
    emitTelemetry(telemetrySink, {
      type: "ui_error_boundary",
      message: errorMessage,
      stack: errorStack,
      diagnostics: officeDiag,
      timestamp: Date.now(),
    });
  }

  function getDiagnosticDetails(): string {
    return JSON.stringify(
      {
        error: errorMessage,
        stack: errorStack,
        office: getOfficeDiagnostics(),
        userAgent:
          typeof navigator !== "undefined" ? navigator.userAgent : undefined,
        url: typeof window !== "undefined" ? window.location.href : undefined,
        timestamp: new Date().toISOString(),
      },
      null,
      2,
    );
  }

  async function copyErrorDetails() {
    try {
      if (typeof navigator !== "undefined" && navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(getDiagnosticDetails());
        copied = true;
        setTimeout(() => {
          copied = false;
        }, 2000);
      }
    } catch (err) {
      console.error("Failed to copy error details:", err);
    }
  }
</script>

<svelte:boundary {onerror}>
  {@render children?.()}

  {#snippet failed(_error, reset)}
    <div
      class="h-screen w-full flex items-center justify-center bg-(--chat-bg) px-4"
      style="font-family: var(--chat-font-mono)"
    >
      <div class="w-full max-w-xl border border-(--chat-border) bg-(--chat-bg-secondary) p-4 space-y-3">
        <div class="text-xs uppercase tracking-widest text-(--chat-text-muted)">
          Unhandled UI Error
        </div>
        <div class="text-sm text-(--chat-text-primary)">
          The chat UI hit an unexpected error.
        </div>
        <pre class="max-h-48 overflow-auto text-xs text-(--chat-error) bg-(--chat-bg) border border-(--chat-border) p-2 whitespace-pre-wrap break-words">
{errorStack || errorMessage}
        </pre>
        <div class="flex gap-2">
          <button
            type="button"
            onclick={() => {
              errorMessage = "";
              errorStack = undefined;
              reset();
            }}
            class="px-3 py-1.5 text-xs border border-(--chat-border) text-(--chat-text-primary) hover:bg-(--chat-bg)"
          >
            Try again
          </button>
          <button
            type="button"
            onclick={copyErrorDetails}
            class="px-3 py-1.5 text-xs border border-(--chat-border) text-(--chat-text-primary) hover:bg-(--chat-bg)"
          >
            {copied ? "Copied!" : "Copy Error Details"}
          </button>
          <button
            type="button"
            onclick={() => window.location.reload()}
            class="px-3 py-1.5 text-xs border border-(--chat-error) text-(--chat-error) hover:bg-(--chat-error) hover:text-(--chat-bg)"
          >
            Reload add-in
          </button>
        </div>
      </div>
    </div>
  {/snippet}
</svelte:boundary>
