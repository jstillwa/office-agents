import { Readability } from "@mozilla/readability";
import TurndownService from "turndown";
import type { FetchProvider, FetchResult, WebContext } from "./types";

function isPrivateIpv4(a: number, b: number, c: number, d: number): boolean {
  // 0.0.0.0/8 (Current network)
  if (a === 0) return true;
  // 10.0.0.0/8 (RFC 1918)
  if (a === 10) return true;
  // 127.0.0.0/8 (Loopback)
  if (a === 127) return true;
  // 169.254.0.0/16 (Link-local / cloud metadata)
  if (a === 169 && b === 254) return true;
  // 172.16.0.0/12 (RFC 1918: 172.16.0.0 - 172.31.255.255)
  if (a === 172 && b >= 16 && b <= 31) return true;
  // 192.168.0.0/16 (RFC 1918)
  if (a === 192 && b === 168) return true;
  // 100.64.0.0/10 (Shared address space / Carrier-grade NAT)
  if (a === 100 && b >= 64 && b <= 127) return true;
  // 192.0.0.0/24 (IETF Protocol Assignments)
  if (a === 192 && b === 0 && c === 0) return true;
  // 192.0.2.0/24 (TEST-NET-1)
  if (a === 192 && b === 0 && c === 2) return true;
  // 198.18.0.0/15 (Network benchmark tests)
  if (a === 198 && (b === 18 || b === 19)) return true;
  // 198.51.100.0/24 (TEST-NET-2)
  if (a === 198 && b === 51 && c === 100) return true;
  // 203.0.113.0/24 (TEST-NET-3)
  if (a === 203 && b === 0 && c === 113) return true;
  // 224.0.0.0/4 (Multicast)
  if (a >= 224 && a <= 239) return true;
  // 240.0.0.0/4 (Reserved)
  if (a >= 240) return true;
  // 255.255.255.255 (Broadcast)
  if (a === 255 && b === 255 && c === 255 && d === 255) return true;
  return false;
}

export function assertSafeUrl(input: string): URL {
  let parsed: URL;
  try {
    parsed = new URL(input);
  } catch {
    throw new Error(`Invalid URL: "${input}"`);
  }

  if (parsed.protocol !== "https:") {
    throw new Error(
      `Insecure protocol "${parsed.protocol}" rejected for URL "${input}". Only HTTPS is allowed.`,
    );
  }

  const hostname = parsed.hostname.toLowerCase();

  // Hostname checks (localhost, local, internal)
  if (
    hostname === "localhost" ||
    hostname.endsWith(".localhost") ||
    hostname.endsWith(".local") ||
    hostname.endsWith(".internal") ||
    hostname === "instance-data"
  ) {
    throw new Error(
      `Access to local/private host "${parsed.hostname}" is prohibited (SSRF guard).`,
    );
  }

  // IPv4 dotted decimal check
  const ipv4Match = hostname.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)$/);
  if (ipv4Match) {
    const a = Number.parseInt(ipv4Match[1], 10);
    const b = Number.parseInt(ipv4Match[2], 10);
    const c = Number.parseInt(ipv4Match[3], 10);
    const d = Number.parseInt(ipv4Match[4], 10);
    if (a > 255 || b > 255 || c > 255 || d > 255 || isPrivateIpv4(a, b, c, d)) {
      throw new Error(
        `Access to private/internal IP "${parsed.hostname}" is prohibited (SSRF guard).`,
      );
    }
  }

  // IPv6 bracketed check
  if (hostname.startsWith("[") && hostname.endsWith("]")) {
    const rawIp6 = hostname.slice(1, -1);
    if (rawIp6 === "::1" || rawIp6 === "::") {
      throw new Error(
        `Access to IPv6 loopback "${parsed.hostname}" is prohibited (SSRF guard).`,
      );
    }
    // Link-local fe80::/10 (fe8, fe9, fea, feb)
    if (/^fe[89ab]/i.test(rawIp6)) {
      throw new Error(
        `Access to IPv6 link-local "${parsed.hostname}" is prohibited (SSRF guard).`,
      );
    }
    // Unique Local Address fc00::/7 (fc, fd)
    if (/^f[cd]/i.test(rawIp6)) {
      throw new Error(
        `Access to IPv6 private address "${parsed.hostname}" is prohibited (SSRF guard).`,
      );
    }
    // IPv4-mapped IPv6: ::ffff:A.B.C.D or ::ffff:hex:hex
    if (rawIp6.startsWith("::ffff:")) {
      const remainder = rawIp6.slice(7);
      const mappedIpv4 = remainder.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)$/);
      if (mappedIpv4) {
        const a = Number.parseInt(mappedIpv4[1], 10);
        const b = Number.parseInt(mappedIpv4[2], 10);
        const c = Number.parseInt(mappedIpv4[3], 10);
        const d = Number.parseInt(mappedIpv4[4], 10);
        if (isPrivateIpv4(a, b, c, d)) {
          throw new Error(
            `Access to private/internal IP "${parsed.hostname}" is prohibited (SSRF guard).`,
          );
        }
      } else {
        const hexParts = remainder.split(":");
        if (hexParts.length === 2) {
          const high = Number.parseInt(hexParts[0], 16);
          const low = Number.parseInt(hexParts[1], 16);
          const a = (high >> 8) & 0xff;
          const b = high & 0xff;
          const c = (low >> 8) & 0xff;
          const d = low & 0xff;
          if (isPrivateIpv4(a, b, c, d)) {
            throw new Error(
              `Access to private/internal IP "${parsed.hostname}" is prohibited (SSRF guard).`,
            );
          }
        }
      }
    }
  }

  return parsed;
}

function isHtmlContentType(contentType: string): boolean {
  const ct = contentType.split(";")[0].trim().toLowerCase();
  return ct === "text/html" || ct === "application/xhtml+xml";
}

function isTextContentType(contentType: string): boolean {
  const ct = contentType.split(";")[0].trim().toLowerCase();
  if (ct.startsWith("text/")) return true;
  return (
    ct === "application/json" ||
    ct === "application/xml" ||
    ct === "application/javascript"
  );
}

function getApiKey(
  context: WebContext,
  providerId: string,
): string | undefined {
  return context.apiKeys?.[providerId];
}

async function fetchWithProxy(
  url: string,
  proxyUrl?: string,
  init?: RequestInit,
): Promise<Response> {
  assertSafeUrl(url);
  if (proxyUrl) {
    assertSafeUrl(proxyUrl);
    try {
      return await fetch(`${proxyUrl}/?url=${encodeURIComponent(url)}`, init);
    } catch (err) {
      throw new Error(
        `CORS proxy fetch failed: ${err instanceof Error ? err.message : err}`,
      );
    }
  }

  try {
    return await fetch(url, init);
  } catch {
    throw new Error(
      "Fetch blocked by CORS and no CORS proxy is configured. Enable the CORS proxy in Settings.",
    );
  }
}

function extractContentFromHtml(url: string, html: string): FetchResult {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const base = doc.createElement("base");
  base.href = url;
  doc.head.prepend(base);

  const reader = new Readability(doc);
  const article = reader.parse();

  const td = new TurndownService({
    headingStyle: "atx",
    codeBlockStyle: "fenced",
    bulletListMarker: "-",
  });
  td.remove(["script", "style", "footer"]);

  let title: string;
  let content: string;
  const metadata: Record<string, string> = { URL: url };

  if (article) {
    title = article.title;
    if (article.byline) metadata.Author = article.byline;
    if (article.siteName) metadata.Site = article.siteName;
    content = td.turndown(article.content).trim();
  } else {
    title = doc.querySelector("title")?.textContent ?? url;
    content = td.turndown(doc.body?.innerHTML ?? "").trim();
  }

  return {
    kind: "text",
    contentType: "text/markdown",
    text: content,
    title,
    metadata,
  };
}

const basicFetchProvider: FetchProvider = {
  id: "basic",
  async fetch(url, context): Promise<FetchResult> {
    const resp = await fetchWithProxy(url, context.proxyUrl);
    if (!resp.ok) {
      throw new Error(`Fetch failed: ${resp.status} ${resp.statusText}`);
    }

    const contentType = resp.headers.get("content-type") ?? "";

    if (isHtmlContentType(contentType)) {
      return extractContentFromHtml(url, await resp.text());
    }

    if (isTextContentType(contentType)) {
      return {
        kind: "text",
        contentType,
        text: await resp.text(),
      };
    }

    return {
      kind: "binary",
      contentType,
      data: new Uint8Array(await resp.arrayBuffer()),
    };
  },
};

type ExaTextOptions =
  | boolean
  | {
      maxCharacters?: number;
      includeHtmlTags?: boolean;
    };

interface ExaContentsRequest {
  urls: string[];
  text?: ExaTextOptions;
}

interface ExaContentsResult {
  id?: string;
  url?: string;
  title?: string;
  text?: string;
  summary?: string;
  highlights?: string[];
  author?: string;
  publishedDate?: string;
}

interface ExaContentsResponse {
  requestId?: string;
  results?: ExaContentsResult[];
}

function toExaTextResult(
  inputUrl: string,
  response: ExaContentsResponse,
): FetchResult | null {
  const first = response.results?.[0];
  if (!first) return null;

  const text =
    first.text?.trim() ||
    first.summary?.trim() ||
    first.highlights?.filter(Boolean).join("\n\n").trim() ||
    "";

  if (!text) return null;

  const metadata: Record<string, string> = {
    URL: first.url || inputUrl,
  };

  if (first.author) metadata.Author = first.author;
  if (first.publishedDate) metadata.Published = first.publishedDate;
  if (response.requestId) metadata.ExaRequestId = response.requestId;
  if (first.id) metadata.ExaId = first.id;

  return {
    kind: "text",
    contentType: "text/markdown",
    title: first.title || first.url || inputUrl,
    text,
    metadata,
  };
}

const exaFetchProvider: FetchProvider = {
  id: "exa",
  requiresApiKey: true,
  async fetch(url, context): Promise<FetchResult> {
    assertSafeUrl(url);
    const apiKey = getApiKey(context, "exa");
    if (!apiKey) {
      throw new Error(
        "Exa fetch requires an API key. Configure it in Settings > Web > API Keys.",
      );
    }

    const endpoint = "https://api.exa.ai/contents";
    const body: ExaContentsRequest = {
      urls: [url],
      text: true,
    };

    const resp = await fetchWithProxy(endpoint, context.proxyUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
      },
      body: JSON.stringify(body),
    });

    if (resp.ok) {
      const json = (await resp.json()) as ExaContentsResponse;
      const parsed = toExaTextResult(url, json);
      if (parsed) return parsed;
      return basicFetchProvider.fetch(url, context);
    }

    if (resp.status === 401 || resp.status === 403) {
      throw new Error(`Exa fetch failed: ${resp.status} ${resp.statusText}`);
    }

    return basicFetchProvider.fetch(url, context);
  },
};

const PROVIDERS: Record<string, FetchProvider> = {
  basic: basicFetchProvider,
  exa: exaFetchProvider,
};

export function listFetchProviders(): string[] {
  return Object.keys(PROVIDERS);
}

export function getFetchProvider(providerId?: string): FetchProvider {
  if (!providerId) return basicFetchProvider;
  return PROVIDERS[providerId] || basicFetchProvider;
}

export async function fetchWeb(
  url: string,
  context: WebContext = {},
  providerId?: string,
): Promise<FetchResult> {
  assertSafeUrl(url);
  const provider = getFetchProvider(providerId);
  return provider.fetch(url, context);
}
