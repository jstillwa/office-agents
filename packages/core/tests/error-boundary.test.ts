// @vitest-environment happy-dom
import type { TelemetryEvent, TelemetrySink } from "@office-agents/sdk";
import { mount, unmount } from "svelte";
import { describe, expect, it, vi } from "vitest";
import ErrorBoundaryWrapper from "./ErrorBoundaryWrapper.svelte";

describe("ErrorBoundary component", () => {
  it("captures render errors, emits diagnostics to telemetrySink, and provides copy button", async () => {
    const emittedEvents: TelemetryEvent[] = [];
    const sink: TelemetrySink = {
      emit: vi.fn((event) => {
        emittedEvents.push(event);
      }),
    };

    // Mock Office diagnostics
    (window as unknown as { Office: unknown }).Office = {
      context: {
        diagnostics: {
          host: "Excel",
          platform: "PC",
          version: "16.0.12345",
        },
      },
    };

    // Mock navigator.clipboard
    let clipboardText = "";
    Object.defineProperty(navigator, "clipboard", {
      value: {
        writeText: vi.fn(async (text: string) => {
          clipboardText = text;
        }),
      },
      configurable: true,
    });

    const container = document.createElement("div");
    document.body.appendChild(container);

    const app = mount(ErrorBoundaryWrapper, {
      target: container,
      props: {
        telemetrySink: sink,
        shouldThrow: true,
      },
    });

    // Wait for boundary to catch error
    await new Promise((resolve) => setTimeout(resolve, 50));

    expect(sink.emit).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "ui_error_boundary",
        message: "Simulated component render failure",
        diagnostics: {
          host: "Excel",
          platform: "PC",
          version: "16.0.12345",
        },
      }),
    );

    const copyButton = Array.from(
      container.querySelectorAll("button"),
    ).find((btn) => btn.textContent?.includes("Copy Error Details"));
    expect(copyButton).toBeDefined();

    await copyButton?.click();
    expect(navigator.clipboard.writeText).toHaveBeenCalled();
    expect(clipboardText).toContain("Simulated component render failure");
    expect(clipboardText).toContain("Excel");

    unmount(app);
    document.body.removeChild(container);
  });
});
