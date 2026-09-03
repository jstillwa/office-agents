<script lang="ts">
  import type { ChatController } from "../src/chat/chat-controller";
  import { setChatContext } from "../src/chat/chat-runtime-context";
  import FilesPanel from "../src/chat/files-panel.svelte";

  let { controller }: { controller: ChatController } = $props();
  setChatContext(controller);
</script>

<FilesPanel />
