import JSZip from "jszip";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mockOfficeHost } from "../../core/tests/helpers/office-mock";
import { parseXml } from "../src/lib/pptx/xml-utils";
import {
  createEditSlideXmlTool,
  MAX_SCRIPT_SIZE,
} from "../src/lib/tools/edit-slide-xml";

describe("edit_slide_xml tool", () => {
  let mockContext: any;
  let agentContext: any;
  let restoreHost: () => void;
  let slideZipBase64: string;

  beforeEach(async () => {
    agentContext = {
      readFile: vi.fn().mockResolvedValue(""),
      readFileBuffer: vi.fn().mockResolvedValue(new Uint8Array()),
      writeFile: vi.fn().mockResolvedValue(undefined),
    };

    const zip = new JSZip();
    zip.file(
      "ppt/slides/slide1.xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Hello World</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
    );
    zip.file(
      "ppt/slides/_rels/slide1.xml.rels",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    );
    slideZipBase64 = await zip.generateAsync({ type: "base64" });

    const mockSlide = {
      id: "slide-id-1",
      exportAsBase64: vi.fn(() => ({ value: slideZipBase64 })),
      delete: vi.fn(),
    };

    mockContext = {
      presentation: {
        slides: {
          items: [mockSlide],
          load: vi.fn(),
          getItemAt: vi.fn(() => mockSlide),
        },
        getSelectedSlides: vi.fn(() => ({
          load: vi.fn(),
          items: [],
        })),
        insertSlidesFromBase64: vi.fn(),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    const host = mockOfficeHost("PowerPoint", mockContext);
    restoreHost = host.restore;
  });

  afterEach(() => {
    restoreHost();
  });

  describe("bounds validation", () => {
    it("rejects negative slide index", async () => {
      const tool = createEditSlideXmlTool(agentContext);
      const result = await tool.execute("call-neg", {
        slide_index: -1,
        code: "return true;",
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain(
        "slide_index must be a non-negative integer",
      );
    });

    it("rejects non-integer slide index", async () => {
      const tool = createEditSlideXmlTool(agentContext);
      const result = await tool.execute("call-float", {
        slide_index: 1.5,
        code: "return true;",
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain(
        "slide_index must be a non-negative integer",
      );
    });

    it("rejects slide index out of range when exceeding total slides", async () => {
      const tool = createEditSlideXmlTool(agentContext);
      const result = await tool.execute("call-oob", {
        slide_index: 5,
        code: "return true;",
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain("Slide index 5 out of range (0-0)");
    });

    it("rejects scripts exceeding MAX_SCRIPT_SIZE", async () => {
      const tool = createEditSlideXmlTool(agentContext);
      const oversizedScript = "x".repeat(MAX_SCRIPT_SIZE + 10);
      const result = await tool.execute("call-oversized", {
        slide_index: 0,
        code: oversizedScript,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain("exceeds maximum limit");
    });
  });

  describe("XML sanitization and safe parsing", () => {
    it("sanitizes unescaped XML ampersands on save", async () => {
      const tool = createEditSlideXmlTool(agentContext);
      const result = await tool.execute("call-sanitize-amp", {
        slide_index: 0,
        code: `
          const xml = await zip.file("ppt/slides/slide1.xml").async("string");
          const modified = xml.replace("Hello World", "AT&T & Sons");
          zip.file("ppt/slides/slide1.xml", modified);
          markDirty();
          return "modified";
        `,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      expect(parsed._modifiedSlide).toBe(0);

      // Verify that insertSlidesFromBase64 was called with sanitized XML
      expect(
        mockContext.presentation.insertSlidesFromBase64,
      ).toHaveBeenCalled();
      const base64Arg =
        mockContext.presentation.insertSlidesFromBase64.mock.calls[0][0];
      const savedZip = await JSZip.loadAsync(base64Arg, { base64: true });
      const savedXml = await savedZip
        .file("ppt/slides/slide1.xml")!
        .async("string");
      expect(savedXml).toContain("AT&amp;T &amp; Sons");
    });

    it("blocks adding external references in slide XML", async () => {
      const tool = createEditSlideXmlTool(agentContext);
      const result = await tool.execute("call-ext-rel", {
        slide_index: 0,
        code: `
          const rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://evil.com" TargetMode="External"/></Relationships>';
          zip.file("ppt/slides/_rels/slide1.xml.rels", rels);
          markDirty();
        `,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain("Adding external references is blocked");
    });

    it("provides safe parseXml helper that parses valid XML and throws on malformed XML", () => {
      const validDoc = parseXml("<root><item id='1'>Test</item></root>");
      expect(validDoc.documentElement.tagName).toBe("root");
      expect(validDoc.getElementsByTagName("item")[0].textContent).toBe("Test");

      expect(() => parseXml("<root><unclosed></root>")).toThrow(
        /XML parse error/,
      );
    });
  });
});
