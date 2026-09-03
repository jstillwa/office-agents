import { compile } from "svelte/compiler";
import { defineConfig } from "vitest/config";

export default defineConfig({
  plugins: [
    {
      name: "svelte-transform",
      transform(code, id) {
        if (id.endsWith(".svelte")) {
          const compiled = compile(code, {
            filename: id,
            generate: "client",
          });
          return {
            code: compiled.js.code,
            map: compiled.js.map,
          };
        }
      },
    },
  ],
  test: {
    include: ["tests/**/*.test.ts"],
    server: {
      deps: {
        inline: ["pdfjs-dist", "lucide-svelte"],
      },
    },
  },
  resolve: {
    conditions: ["browser"],
  },
  define: {
    DOMMatrix: "Object",
  },
});
