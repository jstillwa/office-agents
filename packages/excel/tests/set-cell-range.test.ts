import { beforeEach, describe, expect, it, vi } from "vitest";
import { mockOfficeHost } from "../../core/tests/helpers/office-mock";
import { setCellRangeTool } from "../src/lib/tools/set-cell-range";

describe("set_cell_range tool", () => {
  let mockRange: any;
  let mockSheet: any;
  let mockContext: any;
  let rangesMap: Map<string, any>;

  beforeEach(() => {
    rangesMap = new Map();

    const createMockRange = (
      addr: string,
      rowCount = 1,
      colCount = 1,
      values: any[][] = [[null]],
    ) => {
      const r = {
        address: `Sheet1!${addr}`,
        rowCount,
        columnCount: colCount,
        values,
        formulas: values.map((row) => row.map(() => "")),
        format: {
          load: vi.fn(),
          font: {},
          fill: {},
          borders: { getItem: vi.fn(() => ({})) },
          autofitColumns: vi.fn(),
          autofitRows: vi.fn(),
        },
        load: vi.fn(),
        copyFrom: vi.fn(),
      };
      rangesMap.set(addr, r);
      return r;
    };

    mockRange = createMockRange("A1", 1, 1, [[null]]);

    mockSheet = {
      id: "sheet-guid-1",
      name: "Sheet1",
      load: vi.fn(),
      getRange: vi.fn((addr: string) => {
        if (!rangesMap.has(addr)) {
          // Parse dimensions if range has colon
          if (addr.includes(":")) {
            const parts = addr.split(":");
            const startCol = parts[0].charCodeAt(0) - 65;
            const endCol = parts[1].charCodeAt(0) - 65;
            const startRow = Number.parseInt(parts[0].slice(1), 10);
            const endRow = Number.parseInt(parts[1].slice(1), 10);
            const rows = endRow - startRow + 1;
            const cols = endCol - startCol + 1;
            return createMockRange(
              addr,
              rows,
              cols,
              Array.from({ length: rows }, () => Array(cols).fill(null)),
            );
          }
          return createMockRange(addr);
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

  it("fails with overwrite error when cell is non-empty and allow_overwrite is false", async () => {
    mockRange.values = [["Existing Value"]];

    const result = await setCellRangeTool.execute("call-1", {
      sheetId: 1,
      range: "A1",
      cells: [[{ value: "New Value" }]],
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain("Would overwrite 1 non-empty cell(s)");
    expect(parsed.error).toContain("retry with allow_overwrite set to true");
  });

  it("succeeds when allow_overwrite is true even if cell is non-empty", async () => {
    mockRange.values = [["Existing Value"]];

    const result = await setCellRangeTool.execute("call-2", {
      sheetId: 1,
      range: "A1",
      cells: [[{ value: "Replacement Value" }]],
      allow_overwrite: true,
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(mockRange.values).toEqual([["Replacement Value"]]);
  });

  it("auto-expands 1x1 initial range to 2x3 dimensions when cells array is 2x3", async () => {
    const cells2x3 = [
      [{ value: "A1" }, { value: "B1" }, { value: "C1" }],
      [{ value: "A2" }, { value: "B2" }, { value: "C2" }],
    ];

    const result = await setCellRangeTool.execute("call-3", {
      sheetId: 1,
      range: "A1",
      cells: cells2x3,
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    // Adjusted range A1:C2
    expect(parsed.messages[0]).toContain("Adjusted range from A1 to A1:C2");
    const expandedRange = rangesMap.get("A1:C2");
    expect(expandedRange).toBeDefined();
    expect(expandedRange.values).toEqual([
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"],
    ]);
  });

  it("tracks dirty ranges properly in response", async () => {
    const result = await setCellRangeTool.execute("call-4", {
      sheetId: 1,
      range: "A1:B2",
      cells: [
        [{ value: 1 }, { value: 2 }],
        [{ value: 3 }, { value: 4 }],
      ],
      copyToRange: "A1:D4",
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.success).toBe(true);
    expect(parsed._dirtyRanges).toEqual([
      { sheetId: 1, range: "A1:B2" },
      { sheetId: 1, range: "A1:D4" },
    ]);
  });
});
