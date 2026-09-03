// @vitest-environment happy-dom
import { AgentContext, type RuntimeState } from "@office-agents/sdk";
import { mount, unmount } from "svelte";
import { writable } from "svelte/store";
import { describe, expect, it, vi } from "vitest";
import type { AppAdapter } from "../src/chat/app-adapter";
import type { ChatController } from "../src/chat/chat-controller";
import FilesPanelWrapper from "./FilesPanelWrapper.svelte";

describe("FilesPanel component", () => {
  it("notifies users of deletion failures with an alert banner", async () => {
    const context = new AgentContext({
      namespace: { prefix: "test", documentId: "doc-files" },
    });

    // Create a file in VFS
    await context.writeFile("sample.txt", "hello");

    // Spy on context.deleteFile to simulate a deletion failure
    vi.spyOn(context, "deleteFile").mockRejectedValue(
      new Error("Permission denied: cannot delete file"),
    );

    const stateStore = writable<Partial<RuntimeState>>({
      vfsInvalidatedAt: 1,
    });

    const mockController = {
      context,
      state: stateStore,
      removeUpload: vi.fn(),
    } as unknown as ChatController;

    const container = document.createElement("div");
    document.body.appendChild(container);

    const app = mount(FilesPanelWrapper, {
      target: container,
      props: {
        controller: mockController,
      },
    });

    // Wait for initial files list to load
    await new Promise((resolve) => setTimeout(resolve, 50));

    // Find delete button
    const deleteBtn = container.querySelector(
      'button[title="Delete"]',
    ) as HTMLButtonElement | null;
    expect(deleteBtn).not.toBeNull();

    await deleteBtn?.click();
    await new Promise((resolve) => setTimeout(resolve, 50));

    // Alert banner should appear notifying of failure
    const alertBanner = container.querySelector('[role="alert"]');
    expect(alertBanner).not.toBeNull();
    expect(alertBanner?.textContent).toContain(
      "Permission denied: cannot delete file",
    );

    // Dismissing banner works
    const dismissBtn = alertBanner?.querySelector("button");
    await dismissBtn?.click();
    expect(container.querySelector('[role="alert"]')).toBeNull();

    unmount(app);
    document.body.removeChild(container);
  });
});
