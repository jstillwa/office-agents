<script lang="ts">
  import type { ChatController } from "../src/chat/chat-controller";
  import { setChatContext } from "../src/chat/chat-runtime-context";
  import MessageList from "../src/chat/message-list.svelte";

  let { controller }: { controller: ChatController } = $props();
  setChatContext(controller);
</script>

<MessageList />
