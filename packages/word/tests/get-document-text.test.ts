import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mockOfficeHost } from "../../core/tests/helpers/office-mock";
import { getDocumentTextTool } from "../src/lib/tools/get-document-text";

describe("get_document_text tool", () => {
  let mockContext: any;
  let restoreHost: () => void;
  let paragraphs: any[];

  beforeEach(() => {
    paragraphs = [
      {
        text: "Title Paragraph",
        style: "Title",
        alignment: "Center",
        listItemOrNullObject: { isNullObject: true, load: vi.fn() },
        load: vi.fn(),
      },
      {
        text: "First bullet point",
        style: "List Paragraph",
        alignment: "Left",
        listItemOrNullObject: {
          isNullObject: false,
          level: 0,
          listString: "1.",
          load: vi.fn(),
        },
        load: vi.fn(),
      },
      {
        text: "Sub bullet point",
        style: "List Paragraph",
        alignment: "Left",
        listItemOrNullObject: {
          isNullObject: false,
          level: 1,
          listString: "a.",
          load: vi.fn(),
        },
        load: vi.fn(),
      },
      {
        text: "Closing remarks",
        style: "Normal",
        alignment: "Left",
        listItemOrNullObject: { isNullObject: true, load: vi.fn() },
        load: vi.fn(),
      },
    ];

    mockContext = {
      document: {
        body: {
          paragraphs: {
            items: paragraphs,
            load: vi.fn(),
          },
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    const host = mockOfficeHost("Word", mockContext);
    restoreHost = host.restore;
  });

  afterEach(() => {
    restoreHost();
  });

  it("reads all paragraphs with formatting and list items by default", async () => {
    const result = await getDocumentTextTool.execute("call-all", {});
    const parsed = JSON.parse((result.content[0] as { text: string }).text);

    expect(parsed.totalParagraphs).toBe(4);
    expect(parsed.showing).toEqual({ start: 0, end: 4 });
    expect(parsed.paragraphs.length).toBe(4);

    expect(parsed.paragraphs[0]).toEqual({
      index: 0,
      text: "Title Paragraph",
      style: "Title",
      alignment: "Center",
    });

    expect(parsed.paragraphs[1]).toEqual({
      index: 1,
      text: "First bullet point",
      style: "List Paragraph",
      alignment: "Left",
      listLevel: 0,
      listString: "1.",
    });

    expect(parsed.paragraphs[2]).toEqual({
      index: 2,
      text: "Sub bullet point",
      style: "List Paragraph",
      alignment: "Left",
      listLevel: 1,
      listString: "a.",
    });
  });

  it("slices paragraphs correctly with startParagraph and endParagraph", async () => {
    const result = await getDocumentTextTool.execute("call-slice", {
      startParagraph: 1,
      endParagraph: 3,
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);

    expect(parsed.totalParagraphs).toBe(4);
    expect(parsed.showing).toEqual({ start: 1, end: 3 });
    expect(parsed.paragraphs.length).toBe(2);
    expect(parsed.paragraphs[0].text).toBe("First bullet point");
    expect(parsed.paragraphs[1].text).toBe("Sub bullet point");
  });

  it("omits styles and list info when includeFormatting is false", async () => {
    const result = await getDocumentTextTool.execute("call-no-formatting", {
      includeFormatting: false,
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);

    expect(parsed.totalParagraphs).toBe(4);
    expect(parsed.paragraphs[0]).toEqual({
      index: 0,
      text: "Title Paragraph",
    });
    expect(parsed.paragraphs[1]).toEqual({
      index: 1,
      text: "First bullet point",
    });
    expect(parsed.paragraphs[1].style).toBeUndefined();
    expect(parsed.paragraphs[1].listLevel).toBeUndefined();
  });
});
