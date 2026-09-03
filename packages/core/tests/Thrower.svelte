<script lang="ts">
  let { shouldThrow = false }: { shouldThrow?: boolean } = $props();
  if (shouldThrow) {
    throw new Error("Simulated component render failure");
  }
</script>

<div>Normal render</div>
