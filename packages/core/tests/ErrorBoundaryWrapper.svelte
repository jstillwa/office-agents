<script lang="ts">
  import type { TelemetrySink } from "@office-agents/sdk";
  import ErrorBoundary from "../src/chat/error-boundary.svelte";
  import Thrower from "./Thrower.svelte";

  let {
    telemetrySink,
    shouldThrow = false,
  }: {
    telemetrySink?: TelemetrySink;
    shouldThrow?: boolean;
  } = $props();
</script>

<ErrorBoundary {telemetrySink}>
  <Thrower {shouldThrow} />
</ErrorBoundary>
