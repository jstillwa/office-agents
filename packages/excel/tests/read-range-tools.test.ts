import { beforeEach, describe, expect, it, vi } from "vitest";
import { mockOfficeHost } from "../../core/tests/helpers/office-mock";
import { getCellRangesTool } from "../src/lib/tools/get-cell-ranges";
import { getRangeAsCsvTool } from "../src/lib/tools/get-range-as-csv";

describe("read range tools", () => {
  let mockSheet: any;
  let mockContext: any;
  let rangesMap: Map<string, any>;

  beforeEach(() => {
    rangesMap = new Map();

    const createMockCell = () => ({
      format: {
        font: {
          load: vi.fn(),
          name: "Calibri",
          size: 11,
          bold: true,
          color: "#FF0000",
        },
        fill: {
          load: vi.fn(),
          color: "#00FF00",
        },
      },
    });

    const createMockRange = (
      addr: string,
      rowCount = 1,
      colCount = 1,
      values: any[][] = [[""]],
    ) => {
      const r = {
        address: `Sheet1!${addr}`,
        rowCount,
        columnCount: colCount,
        values,
        formulas: values.map((row) => row.map(() => "")),
        getCell: vi.fn(() => createMockCell()),
        load: vi.fn(),
      };
      rangesMap.set(addr, r);
      return r;
    };

    mockSheet = {
      id: "sheet-guid-1",
      name: "Sheet1",
      load: vi.fn(),
      getUsedRangeOrNullObject: vi.fn(() => ({
        isNullObject: false,
        address: "Sheet1!A1:C10",
        load: vi.fn(),
      })),
      getRange: vi.fn((addr: string) => {
        if (!rangesMap.has(addr)) {
          return createMockRange(addr, 2, 2, [
            ["Item", "Price"],
            ["Apple", 1.5],
          ]);
        }
        return rangesMap.get(addr);
      }),
    };

    mockContext = {
      workbook: {
        worksheets: {
          items: [mockSheet],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    mockOfficeHost("Excel", mockContext);
  });

  describe("get_cell_ranges", () => {
    it("retrieves cell values, styles, and dimensions", async () => {
      rangesMap.set("A1:B2", {
        address: "Sheet1!A1:B2",
        rowCount: 2,
        columnCount: 2,
        values: [
          ["Header 1", "Header 2"],
          ["Data 1", 42],
        ],
        formulas: [
          ["", ""],
          ["", "=SUM(1, 41)"],
        ],
        getCell: vi.fn(() => ({
          format: {
            font: {
              load: vi.fn(),
              bold: true,
              name: "Calibri",
              size: 12,
              color: "#112233",
            },
            fill: { load: vi.fn(), color: "#445566" },
          },
        })),
        load: vi.fn(),
      });

      const result = await getCellRangesTool.execute("call-read-1", {
        sheetId: 1,
        ranges: ["A1:B2"],
        includeStyles: true,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      expect(parsed.worksheet.cells.A1).toBe("Header 1");
      expect(parsed.worksheet.cells.B2).toBe(42);
      expect(parsed.worksheet.formulas.B2).toBe("=SUM(1, 41)");
      expect(parsed.worksheet.styles.A1).toEqual({
        bold: true,
        family: "Calibri",
        sz: 12,
        color: "#112233",
        fgColor: "#445566",
      });
      expect(parsed.hasMore).toBe(false);
    });

    it("respects cell limit and sets hasMore flag", async () => {
      rangesMap.set("A1:B2", {
        address: "Sheet1!A1:B2",
        rowCount: 2,
        columnCount: 2,
        values: [
          ["V1", "V2"],
          ["V3", "V4"],
        ],
        formulas: [
          ["", ""],
          ["", ""],
        ],
        getCell: vi.fn(() => ({
          format: { font: { load: vi.fn() }, fill: { load: vi.fn() } },
        })),
        load: vi.fn(),
      });

      const result = await getCellRangesTool.execute("call-read-limit", {
        sheetId: 1,
        ranges: ["A1:B2", "C1:D2"],
        cellLimit: 2,
        includeStyles: false,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      expect(Object.keys(parsed.worksheet.cells).length).toBe(2);
      expect(parsed.hasMore).toBe(true);
    });
  });

  describe("get_range_as_csv", () => {
    it("returns formatted CSV with headers", async () => {
      rangesMap.set("A1:C2", {
        address: "Sheet1!A1:C2",
        rowCount: 2,
        columnCount: 3,
        values: [
          ["Product", "Quantity", "Price"],
          ["Widget", 10, 19.99],
        ],
        load: vi.fn(),
      });

      const result = await getRangeAsCsvTool.execute("call-csv-1", {
        sheetId: 1,
        range: "A1:C2",
        includeHeaders: true,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      expect(parsed.csv).toBe("Product,Quantity,Price\nWidget,10,19.99");
      expect(parsed.rowCount).toBe(2);
      expect(parsed.columnCount).toBe(3);
    });

    it("excludes headers when includeHeaders is false", async () => {
      rangesMap.set("A1:B2", {
        address: "Sheet1!A1:B2",
        rowCount: 2,
        columnCount: 2,
        values: [
          ["Header 1", "Header 2"],
          ["Row 1", "Row 2"],
        ],
        load: vi.fn(),
      });

      const result = await getRangeAsCsvTool.execute("call-csv-no-header", {
        sheetId: 1,
        range: "A1:B2",
        includeHeaders: false,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      expect(parsed.csv).toBe("Row 1,Row 2");
      expect(parsed.rowCount).toBe(1);
    });

    it("properly escapes commas, quotes, and newlines in CSV cells", async () => {
      rangesMap.set("A1:C2", {
        address: "Sheet1!A1:C2",
        rowCount: 2,
        columnCount: 3,
        values: [
          ['Name, with "quotes"', "Line\nBreak", "Regular"],
          ['Item "1", special', "Second\nline", "Simple"],
        ],
        load: vi.fn(),
      });

      const result = await getRangeAsCsvTool.execute("call-csv-escape", {
        sheetId: 1,
        range: "A1:C2",
        includeHeaders: true,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      // Double quotes inside quoted fields
      expect(parsed.csv).toContain('"Name, with ""quotes"""');
      expect(parsed.csv).toContain('"Line\nBreak"');
      expect(parsed.csv).toContain('"Item ""1"", special"');
    });

    it("handles maxRows pagination and hasMore flag", async () => {
      rangesMap.set("A1:A5", {
        address: "Sheet1!A1:A5",
        rowCount: 5,
        columnCount: 1,
        values: [["H"], ["R1"], ["R2"], ["R3"], ["R4"]],
        load: vi.fn(),
      });

      const result = await getRangeAsCsvTool.execute("call-csv-maxrows", {
        sheetId: 1,
        range: "A1:A5",
        includeHeaders: true,
        maxRows: 2,
      });

      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      expect(parsed.rowCount).toBe(2);
      expect(parsed.hasMore).toBe(true);
      expect(parsed.csv).toBe("H\nR1");
    });
  });
});
