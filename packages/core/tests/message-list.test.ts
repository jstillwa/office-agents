// @vitest-environment happy-dom
import type { ChatMessage, MessagePart, RuntimeState } from "@office-agents/sdk";
import { mount, unmount } from "svelte";
import { writable } from "svelte/store";
import { describe, expect, it } from "vitest";
import type { AppAdapter } from "../src/chat/app-adapter";
import type { ChatController } from "../src/chat/chat-controller";
import MessageListWrapper from "./MessageListWrapper.svelte";

describe("MessageList component", () => {
  it("renders fallback indicator without throwing on unknown MessagePart union types", async () => {
    const unknownPart = {
      type: "future_audio_or_binary_payload",
      customData: 12345,
    } as unknown as MessagePart;

    const testMessages: ChatMessage[] = [
      {
        id: "msg-1",
        role: "user",
        timestamp: Date.now(),
        parts: [{ type: "text", text: "Hello with standard text" }],
      },
      {
        id: "msg-2",
        role: "assistant",
        timestamp: Date.now(),
        parts: [
          { type: "text", text: "Standard response text" },
          unknownPart,
        ],
      },
    ];

    const stateStore = writable<Partial<RuntimeState>>({
      messages: testMessages,
      isStreaming: false,
    });

    const mockAdapter: Partial<AppAdapter> = {
      emptyStateMessage: "No messages yet",
      handleLinkClick: () => "default",
    };

    const mockController = {
      state: stateStore,
      adapter: mockAdapter as AppAdapter,
    } as unknown as ChatController;

    const container = document.createElement("div");
    document.body.appendChild(container);

    let app: ReturnType<typeof mount> | undefined;
    expect(() => {
      app = mount(MessageListWrapper, {
        target: container,
        props: {
          controller: mockController,
        },
      });
    }).not.toThrow();

    const fallbackBadge = container.querySelector(
      '[data-testid="unsupported-part-fallback"]',
    );
    expect(fallbackBadge).not.toBeNull();
    expect(fallbackBadge?.textContent).toContain("future_audio_or_binary_payload");

    if (app) {
      unmount(app);
    }
    document.body.removeChild(container);
  });
});
