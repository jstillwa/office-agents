import { describe, expect, it } from "vitest";
import { assertSafeMcpUrl } from "../src/mcp";
import { assertSafeUrl, fetchWeb } from "../src/web/fetch";

describe("assertSafeUrl SSRF guard", () => {
  it("rejects http: URLs", () => {
    expect(() => assertSafeUrl("http://example.com")).toThrow(/Only HTTPS is allowed/);
    expect(() => assertSafeUrl("http://127.0.0.1")).toThrow(/Only HTTPS is allowed/);
    expect(() => assertSafeUrl("http://localhost")).toThrow(/Only HTTPS is allowed/);
    expect(() => assertSafeUrl("http://10.0.0.1")).toThrow(/Only HTTPS is allowed/);
    expect(() => assertSafeUrl("ftp://example.com")).toThrow(/Only HTTPS is allowed/);
  });

  it("rejects 127.0.0.1 and loopback addresses", () => {
    expect(() => assertSafeUrl("https://127.0.0.1")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://127.0.0.1:8080/path")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://127.1.2.3")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://[::1]")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://[::ffff:127.0.0.1]")).toThrow(/SSRF guard/);
  });

  it("rejects localhost", () => {
    expect(() => assertSafeUrl("https://localhost")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://localhost:3000")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://sub.localhost")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://service.local")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://metadata.google.internal")).toThrow(/SSRF guard/);
  });

  it("rejects 169.254.169.254 and link-local cloud metadata addresses", () => {
    expect(() => assertSafeUrl("https://169.254.169.254")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://169.254.169.254/latest/meta-data")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://169.254.1.1")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://[::ffff:169.254.169.254]")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://[fe80::1]")).toThrow(/SSRF guard/);
  });

  it("rejects RFC 1918 private network ranges (10.x, 172.16.x - 172.31.x, 192.168.x)", () => {
    // 10.0.0.0/8
    expect(() => assertSafeUrl("https://10.0.0.1")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://10.254.1.1")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://10.255.255.255")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://[::ffff:10.0.0.1]")).toThrow(/SSRF guard/);

    // 172.16.0.0/12
    expect(() => assertSafeUrl("https://172.16.0.1")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://172.20.10.5")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://172.31.255.255")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://[::ffff:172.16.0.1]")).toThrow(/SSRF guard/);

    // 192.168.0.0/16
    expect(() => assertSafeUrl("https://192.168.0.1")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://192.168.1.1")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://192.168.254.254")).toThrow(/SSRF guard/);
    expect(() => assertSafeUrl("https://[::ffff:192.168.1.1]")).toThrow(/SSRF guard/);
  });

  it("allows safe public HTTPS URLs", () => {
    const url1 = assertSafeUrl("https://example.com/data.json");
    expect(url1.hostname).toBe("example.com");

    const url2 = assertSafeUrl("https://en.wikipedia.org/wiki/Office");
    expect(url2.hostname).toBe("en.wikipedia.org");

    const url3 = assertSafeUrl("https://1.1.1.1/dns-query");
    expect(url3.hostname).toBe("1.1.1.1");
  });

  it("fetchWeb rejects SSRF URLs immediately before attempting network requests", async () => {
    await expect(fetchWeb("http://example.com")).rejects.toThrow(/Only HTTPS is allowed/);
    await expect(fetchWeb("https://127.0.0.1:4017/secret")).rejects.toThrow(/SSRF guard/);
    await expect(fetchWeb("https://169.254.169.254/metadata")).rejects.toThrow(/SSRF guard/);
    await expect(fetchWeb("https://192.168.1.1/admin")).rejects.toThrow(/SSRF guard/);
  });
});

describe("assertSafeMcpUrl", () => {
  it("allows https endpoints anywhere", () => {
    expect(() => assertSafeMcpUrl("https://mcp.internal.example.com/sse")).not.toThrow();
    expect(() => assertSafeMcpUrl("https://localhost:8080/mcp")).not.toThrow();
  });

  it("allows http endpoints only for localhost or 127.0.0.1", () => {
    expect(() => assertSafeMcpUrl("http://localhost:8080/mcp")).not.toThrow();
    expect(() => assertSafeMcpUrl("http://127.0.0.1:8080/mcp")).not.toThrow();
  });

  it("rejects non-https endpoints when host is not localhost/127.0.0.1", () => {
    expect(() => assertSafeMcpUrl("http://example.com/mcp")).toThrow(/Non-HTTPS MCP endpoints are only permitted/);
    expect(() => assertSafeMcpUrl("http://10.0.0.5:8080/mcp")).toThrow(/Non-HTTPS MCP endpoints are only permitted/);
    expect(() => assertSafeMcpUrl("http://192.168.1.50:3000/sse")).toThrow(/Non-HTTPS MCP endpoints are only permitted/);
  });
});
