import { svelte } from "@sveltejs/vite-plugin-svelte";
import tailwindcss from "@tailwindcss/postcss";
import autoprefixer from "autoprefixer";
import { execSync } from "child_process";
import fs from "fs";
import { createRequire } from "module";
import path from "path";
import { fileURLToPath } from "url";
import { defineConfig, type Plugin } from "vite";
import { nodePolyfills } from "vite-plugin-node-polyfills";
import { viteStaticCopy } from "vite-plugin-static-copy";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const pkg = require("./package.json");

function getGitCommit(): string {
  try {
    return execSync("git rev-parse --short HEAD", { encoding: "utf8" }).trim();
  } catch {
    return process.env.GIT_COMMIT || "unknown";
  }
}

function buildCspHeader(): string {
  const gatewayUrl = process.env.VITE_GATEWAY_URL?.trim();
  const mcpHosts = process.env.VITE_MCP_HOSTS?.trim();
  const extraConnect = [gatewayUrl, mcpHosts].filter(Boolean).join(" ");
  const connectSrc = [
    "'self'",
    "https://appsforoffice.microsoft.com",
    "https://*.office.com",
    "https://*.officeapps.live.com",
    extraConnect,
  ]
    .filter(Boolean)
    .join(" ");

  return [
    "frame-ancestors 'self' https://*.office.com https://*.officeapps.live.com https://*.sharepoint.com https://*.office365.com",
    "default-src 'self'",
    "script-src 'self' https://appsforoffice.microsoft.com",
    `connect-src ${connectSrc}`,
    "object-src 'none'",
  ].join("; ");
}

function generateHeadersContent(): string {
  const csp = buildCspHeader();
  return `/*
  Content-Security-Policy: ${csp}
  X-Content-Type-Options: nosniff
  Referrer-Policy: strict-origin-when-cross-origin
`;
}

function headersPlugin(outDir: string): Plugin {
  return {
    name: "generate-headers",
    generateBundle() {
      this.emitFile({
        type: "asset",
        fileName: "_headers",
        source: generateHeadersContent(),
      });
    },
    closeBundle() {
      try {
        fs.writeFileSync(
          path.resolve(outDir, "_headers"),
          generateHeadersContent(),
          "utf8",
        );
      } catch {}
    },
  };
}

async function getHttpsOptions() {
  try {
    const devCerts = await import("office-addin-dev-certs");
    const certs = await devCerts.getHttpsServerOptions();
    return { ca: certs.ca, key: certs.key, cert: certs.cert };
  } catch {
    console.warn("Could not load office-addin-dev-certs, HTTPS disabled");
    return undefined;
  }
}

export default defineConfig(async ({ mode }) => {
  const dev = mode === "development";
  const defaultDevUrl = "https://localhost:3002/";
  const defaultProdUrl = "https://openword.pages.dev/";

  const rawDeployUrl = process.env.VITE_DEPLOY_URL || defaultProdUrl;
  const deployUrl = rawDeployUrl.endsWith("/")
    ? rawDeployUrl
    : `${rawDeployUrl}/`;
  const deployBaseUrl = deployUrl.replace(/\/$/, "");
  const defaultProdBaseUrl = defaultProdUrl.replace(/\/$/, "");
  const defaultDevBaseUrl = defaultDevUrl.replace(/\/$/, "");
  const deployHost = new URL(deployUrl).host;

  const buildInfo = {
    version: pkg.version as string,
    commit: getGitCommit(),
    timestamp: new Date().toISOString(),
    mode,
  };

  try {
    fs.writeFileSync(
      path.resolve(__dirname, "public/_headers"),
      generateHeadersContent(),
      "utf8",
    );
  } catch {}

  return {
    root: "src",
    publicDir: "../public",

    build: {
      outDir: "../dist",
      emptyOutDir: true,
      sourcemap: true,
      rollupOptions: {
        input: {
          taskpane: path.resolve(__dirname, "src/taskpane.html"),
          commands: path.resolve(__dirname, "src/commands.html"),
        },
      },
    },

    resolve: {
      alias: {
        "node:util/types": path.resolve(
          __dirname,
          "src/shims/util-types-shim.js",
        ),
      },
    },

    define: {
      "process.env": JSON.stringify({}),
      "process.versions": "undefined",
      "process.browser": JSON.stringify(true),
      __APP_VERSION__: JSON.stringify(pkg.version),
      __BUILD_INFO__: JSON.stringify(buildInfo),
    },

    css: {
      postcss: {
        plugins: [tailwindcss(), autoprefixer()],
      },
    },

    plugins: [
      svelte(),

      nodePolyfills({
        include: [
          "buffer",
          "stream",
          "util",
          "url",
          "http",
          "https",
          "zlib",
          "path",
          "os",
          "assert",
          "events",
          "querystring",
          "punycode",
          "string_decoder",
          "constants",
          "vm",
          "process",
        ],
        globals: {
          Buffer: true,
          process: true,
        },
      }),

      headersPlugin(path.resolve(__dirname, "dist")),

      viteStaticCopy({
        targets: [
          {
            src: "../public/assets/*",
            dest: "assets",
          },
          ...(dev
            ? [
                {
                  src: "../manifest.xml",
                  dest: ".",
                  transform: {
                    encoding: "utf8" as const,
                    handler(content: string) {
                      if (!process.env.VITE_DEPLOY_URL) return content;
                      return content
                        .replace(new RegExp(defaultDevUrl, "g"), deployUrl)
                        .replace(
                          new RegExp(defaultDevBaseUrl, "g"),
                          deployBaseUrl,
                        );
                    },
                  },
                },
              ]
            : [
                {
                  src: "../manifest.prod.xml",
                  dest: ".",
                  rename: "manifest.xml",
                  transform: {
                    encoding: "utf8" as const,
                    handler(content: string) {
                      let transformed = content
                        .replace(new RegExp(defaultProdUrl, "g"), deployUrl)
                        .replace(
                          new RegExp(defaultProdBaseUrl, "g"),
                          deployBaseUrl,
                        )
                        .replace(/{HOST}/g, deployHost);
                      if (process.env.VITE_ENTRA_CLIENT_ID) {
                        transformed = transformed.replace(
                          /{CLIENT_ID}/g,
                          process.env.VITE_ENTRA_CLIENT_ID,
                        );
                      }
                      return transformed;
                    },
                  },
                },
              ]),
        ],
      }),
    ],

    server: {
      https: await getHttpsOptions(),
      port: 3002,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
  };
});
