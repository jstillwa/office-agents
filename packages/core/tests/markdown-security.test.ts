// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";
import {
  renderMarkdown,
  renderMarkdownSync,
  sanitizeRenderedHtml,
} from "../src/chat/markdown";

describe("Markdown Security & Sanitization", () => {
  describe("DOMPurify hook sanitization", () => {
    it("adds target='_blank' and rel='noopener noreferrer' to external https links", () => {
      const input = '<a href="https://example.com/docs">Documentation</a>';
      const output = sanitizeRenderedHtml(input);
      expect(output).toContain('target="_blank"');
      expect(output).toContain('rel="noopener noreferrer"');
      expect(output).toContain('href="https://example.com/docs"');
    });

    it("does not add target or rel to relative fragment (#) links", () => {
      const input = '<a href="#section-2">Section 2</a>';
      const output = sanitizeRenderedHtml(input);
      expect(output).toContain('href="#section-2"');
      expect(output).not.toContain("target=");
      expect(output).not.toContain("rel=");
    });

    it("handles links rendered via renderMarkdownSync and renderMarkdown", async () => {
      const md = "[External](https://example.com) and [Internal](#heading)";
      const syncResult = renderMarkdownSync(md);
      expect(syncResult).toContain('href="https://example.com"');
      expect(syncResult).toContain('target="_blank"');
      expect(syncResult).toContain('rel="noopener noreferrer"');
      expect(syncResult).toContain('href="#heading"');
      expect(syncResult).not.toMatch(/href="#heading"[^>]*target="_blank"/);

      const asyncResult = await renderMarkdown(md);
      expect(asyncResult).toContain('href="https://example.com"');
      expect(asyncResult).toContain('target="_blank"');
      expect(asyncResult).toContain('rel="noopener noreferrer"');
      expect(asyncResult).toContain('href="#heading"');
    });
  });

  describe("Code block HTML sanitization", () => {
    it("escapes raw HTML in plain code blocks", () => {
      const markdown = '```\n<script>alert("xss")</script>\n```';
      const result = renderMarkdownSync(markdown);
      expect(result).not.toContain("<script>");
      expect(result).toContain("&lt;script&gt;alert(\"xss\")&lt;/script&gt;");

      // Also verify preferPlainCodeBlocks escaping
      const plainResult = renderMarkdownSync(markdown, { preferPlainCodeBlocks: true });
      expect(plainResult).not.toContain("<script>");
      expect(plainResult).toContain("&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;");
    });

    it("sanitizes highlighted code block output and strips dangerous attributes", async () => {
      const markdown = '```javascript\nconst x = 42;\n```';
      const result = await renderMarkdown(markdown);
      expect(result).toContain("<pre");
      expect(result).toContain("42");
      expect(result).toContain("</code></pre>");

      // Malicious injected attributes in raw code block HTML are stripped by sanitizeRenderedHtml
      const injected = '<pre class="shiki"><code onload="alert(1)"><span onerror="alert(2)">code</span></code></pre>';
      const sanitized = sanitizeRenderedHtml(injected);
      expect(sanitized).not.toContain("onload=");
      expect(sanitized).not.toContain("onerror=");
      expect(sanitized).toContain("code");
    });
  });

  describe("Neutralization of unsafe URI protocols & beacons", () => {
    it("neutralizes javascript: URIs", () => {
      const input = '<a href="javascript:alert(1)">click me</a>';
      const output = sanitizeRenderedHtml(input);
      expect(output).not.toContain("javascript:");
      expect(output).not.toContain('href="javascript:alert(1)"');
    });

    it("neutralizes http: URIs when restricted to https: and fragments", () => {
      const input = '<a href="http://insecure.example.com">insecure</a>';
      const output = sanitizeRenderedHtml(input);
      expect(output).not.toContain("http://insecure.example.com");
    });

    it("neutralizes data: and mailto: URIs", () => {
      const input = '<a href="data:text/html,<script>alert(1)</script>">data</a><a href="mailto:test@example.com">mail</a>';
      const output = sanitizeRenderedHtml(input);
      expect(output).not.toContain("data:text/html");
      expect(output).not.toContain("mailto:test@example.com");
    });

    it("neutralizes remote image tracking beacons in markdown and html", async () => {
      const htmlInput = '<img src="https://tracker.com/beacon.png" alt="beacon">';
      const sanitized = sanitizeRenderedHtml(htmlInput);
      expect(sanitized).not.toContain("<img");
      expect(sanitized).not.toContain("https://tracker.com/beacon.png");

      const markdownImage = "![beacon](https://tracker.com/pixel.gif)";
      const syncResult = renderMarkdownSync(markdownImage);
      expect(syncResult).not.toContain("<img");
      expect(syncResult).not.toContain("https://tracker.com/pixel.gif");

      const asyncResult = await renderMarkdown(markdownImage);
      expect(asyncResult).not.toContain("<img");
      expect(asyncResult).not.toContain("https://tracker.com/pixel.gif");
    });
  });
});
