// @vitest-environment happy-dom
import { AgentContext, type TelemetryEvent, type TelemetrySink } from "@office-agents/sdk";
import { describe, expect, it, vi } from "vitest";
import type { AppAdapter } from "../src/chat/app-adapter";
import { ChatController } from "../src/chat/chat-controller";

function createMockAdapter(overrides: Partial<AppAdapter> = {}): AppAdapter {
  return {
    appName: "TestApp",
    appVersion: "1.0.0",
    tools: [],
    buildSystemPrompt: () => "mock system prompt",
    getDocumentId: vi.fn().mockResolvedValue("test-doc-123"),
    ...overrides,
  };
}

function createMockContext(): AgentContext {
  return new AgentContext({
    namespace: { prefix: "test", documentId: "test-doc-123" },
  });
}

describe("ChatController", () => {
  it("attaches rejection handler to runtime.init() and records error", async () => {
    const emittedEvents: TelemetryEvent[] = [];
    const sink: TelemetrySink = {
      emit: vi.fn((event) => {
        emittedEvents.push(event);
      }),
    };

    const failingAdapter = createMockAdapter({
      getDocumentId: vi.fn().mockRejectedValue(new Error("VFS initialization failed")),
      telemetrySink: sink,
    });

    const context = createMockContext();

    // Instantiating ChatController must not throw unhandled rejection
    const controller = new ChatController(failingAdapter, context);

    // Wait for microtask / promise queue to settle
    await new Promise((resolve) => setTimeout(resolve, 50));

    expect(controller.snapshot.error).toBe("VFS initialization failed");
    expect(controller.recentErrors.length).toBeGreaterThanOrEqual(1);
    expect(controller.recentErrors[0].message).toBe("VFS initialization failed");

    expect(sink.emit).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "uncaught_error",
        message: "VFS initialization failed",
      }),
    );

    controller.dispose();
  });

  it("handles window error and unhandledrejection events", () => {
    const emittedEvents: TelemetryEvent[] = [];
    const sink: TelemetrySink = {
      emit: vi.fn((event) => {
        emittedEvents.push(event);
      }),
    };

    const adapter = createMockAdapter({ telemetrySink: sink });
    const context = createMockContext();
    const controller = new ChatController(adapter, context);

    // Trigger window error
    window.dispatchEvent(
      new ErrorEvent("error", {
        message: "Window uncaught exception",
        error: new Error("Window uncaught exception"),
      }),
    );

    expect(controller.snapshot.error).toBe("Window uncaught exception");
    expect(controller.recentErrors.some((e) => e.message === "Window uncaught exception")).toBe(true);

    // Trigger unhandledrejection
    const rejectionEvent = new Event("unhandledrejection") as PromiseRejectionEvent;
    Object.defineProperty(rejectionEvent, "reason", {
      value: new Error("Async promise boom"),
    });
    window.dispatchEvent(rejectionEvent);

    expect(controller.snapshot.error).toBe("Async promise boom");
    expect(controller.recentErrors.some((e) => e.message === "Async promise boom")).toBe(true);

    controller.dispose();
  });

  it("exports diagnostics aggregating host platform, session stats, and recent errors", () => {
    const adapter = createMockAdapter({
      appName: "ExcelAgents",
      appVersion: "2.4.0",
    });
    const context = createMockContext();
    const controller = new ChatController(adapter, context);

    // Record an error
    controller.handleUncaughtError(new Error("Diagnostic sample error"));

    const bundle = controller.exportDiagnostics();
    expect(bundle).toHaveProperty("timestamp");
    expect(bundle.hostPlatform.appName).toBe("ExcelAgents");
    expect(bundle.hostPlatform.appVersion).toBe("2.4.0");
    expect(bundle.hostPlatform).toHaveProperty("userAgent");
    expect(bundle.sessionStats).toHaveProperty("messageCount");
    expect(bundle.sessionStats).toHaveProperty("stats");
    expect(bundle.recentErrors.length).toBe(1);
    expect(bundle.recentErrors[0].message).toBe("Diagnostic sample error");

    controller.dispose();
  });
});
