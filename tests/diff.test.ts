import { describe, expect, it } from "vitest";
import {
  DEFAULT_TOGGLES,
  diffColumns,
  diffSheets,
  normalizeKey,
  suggestKeyColumn,
  valuesDiffer,
  type SheetTable,
} from "../src/diff";

function table(headers: string[], rows: string[][]): SheetTable {
  return { headers, rows };
}

const A = table(
  ["id", "name", "plan"],
  [
    ["1", "Ada", "pro"],
    ["2", "Grace", "basic"],
    ["3", "Alan", "basic"],
  ],
);

describe("diffSheets — row matching by key", () => {
  it("reports identical files as identical", () => {
    const result = diffSheets(A, table(A.headers, A.rows.map((row) => row.slice())), { keyColumn: "id" });
    expect(result.mode).toBe("key");
    expect(result.summary.addedRows).toBe(0);
    expect(result.summary.removedRows).toBe(0);
    expect(result.summary.changedCells).toBe(0);
    expect(result.summary.matchedRows).toBe(3);
  });

  it("reports an added row with its B record number", () => {
    const b = table(A.headers, [...A.rows.map((row) => row.slice()), ["4", "Katherine", "pro"]]);
    const result = diffSheets(A, b, { keyColumn: "id" });
    expect(result.added).toHaveLength(1);
    expect(result.added[0].key).toBe("4");
    expect(result.added[0].record).toBe(5); // header is record 1
    expect(result.summary.addedRows).toBe(1);
    expect(result.summary.removedRows).toBe(0);
  });

  it("reports a removed row with its A record number", () => {
    const b = table(A.headers, [A.rows[0].slice(), A.rows[2].slice()]);
    const result = diffSheets(A, b, { keyColumn: "id" });
    expect(result.removed).toHaveLength(1);
    expect(result.removed[0].key).toBe("2");
    expect(result.removed[0].record).toBe(3);
  });

  it("reports changed cells with both raw values and both record numbers", () => {
    const b = table(A.headers, [["1", "Ada", "team"], A.rows[1].slice(), A.rows[2].slice()]);
    const result = diffSheets(A, b, { keyColumn: "id" });
    expect(result.changedCells).toEqual([
      {
        key: "1",
        aRecord: 2,
        bRecord: 2,
        column: "plan",
        aColumn: 3,
        bColumn: 3,
        oldValue: "pro",
        newValue: "team",
      },
    ]);
    expect(result.summary.changedRows).toBe(1);
    expect(result.summary.changedCells).toBe(1);
  });

  it("distinguishes a missing cell from an empty string", () => {
    // File A's row is short, so the note cell is missing; file B has it as empty.
    const a = table(["id", "note"], [["1"]]);
    const b = table(["id", "note"], [["1", ""]]);
    expect(diffSheets(a, b, { keyColumn: "id" }).changedCells).toHaveLength(1);

    const toggled = diffSheets(a, b, { keyColumn: "id", toggles: { emptyEqualsMissing: true } });
    expect(toggled.changedCells).toHaveLength(0);

    // An empty value is still different from real content, even with the toggle on.
    const c = table(["id", "note"], [["1", "hello"]]);
    expect(
      diffSheets(c, b, { keyColumn: "id", toggles: { emptyEqualsMissing: true } }).changedCells,
    ).toHaveLength(1);
  });

  it("flags duplicate keys and excludes them from comparison", () => {
    const a = table(["id", "name"], [["1", "Ada"], ["1", "Grace"]]);
    const b = table(["id", "name"], [["1", "Ada"]]);
    const result = diffSheets(a, b, { keyColumn: "id" });
    expect(result.ambiguousKeys).toEqual([
      { key: "1", aRecords: [2, 3], bRecords: [2], reason: "duplicate-in-a" },
    ]);
    expect(result.summary.ambiguousKeys).toBe(1);
    expect(result.added).toHaveLength(0);
    expect(result.removed).toHaveLength(0);
    expect(result.changedCells).toHaveLength(0);
    const finding = result.warnings.find((warning) => warning.id === "ambiguous-keys");
    expect(finding?.details.join(" ")).toContain("excluded from cell comparison");
  });

  it("warns and does nothing when the key column is missing from file B", () => {
    const b = table(["name", "plan"], [["Ada", "pro"]]);
    const result = diffSheets(A, b, { keyColumn: "id" });
    expect(result.mode).toBe("none");
    expect(result.key.problem).toContain("no matching column in file B");
    expect(result.warnings.map((warning) => warning.id)).toContain("key-missing");
    expect(result.changedCells).toHaveLength(0);
  });

  it("reports rows with an empty key instead of matching them", () => {
    const a = table(["id", "name"], [["", "Ada"], ["2", "Grace"]]);
    const b = table(["id", "name"], [["", "Ada"], ["2", "Grace"]]);
    const result = diffSheets(a, b, { keyColumn: "id" });
    expect(result.emptyKey).toEqual({ a: [2], b: [2] });
    expect(result.summary.emptyKeyRowsA).toBe(1);
    expect(result.warnings.map((warning) => warning.id)).toContain("empty-keys");
  });

  it("notes rows that moved without changing", () => {
    const b = table(A.headers, [A.rows[1].slice(), A.rows[0].slice(), A.rows[2].slice()]);
    const result = diffSheets(A, b, { keyColumn: "id" });
    expect(result.moved).toEqual([
      { key: "1", aRecord: 2, bRecord: 3 },
      { key: "2", aRecord: 3, bRecord: 2 },
    ]);
    expect(result.changedCells).toHaveLength(0);
    expect(result.notes.map((note) => note.id)).toContain("moved-rows");
  });

  it("suggests a swap when almost everything looks removed", () => {
    const headers = ["id", "name"];
    const rowsA = Array.from({ length: 12 }, (_, index) => [String(index + 1), `Row ${index + 1}`]);
    const a = table(headers, rowsA);
    const b = table(headers, [["99", "Only one"]]);
    const result = diffSheets(a, b, { keyColumn: "id" });
    expect(result.removed).toHaveLength(12);
    expect(result.notes.map((note) => note.id)).toContain("swapped-hint");
  });

  it("does not cry swap on a small file where rows were simply deleted", () => {
    const b = table(A.headers, [A.rows[2].slice()]);
    const result = diffSheets(A, b, { keyColumn: "id" });
    expect(result.removed).toHaveLength(2);
    expect(result.notes.map((note) => note.id)).not.toContain("swapped-hint");
  });
});

describe("diffSheets — positional fallback", () => {
  it("compares rows by position when explicitly requested", () => {
    const a = table(["id", "name"], [["1", "Ada"], ["2", "Grace"]]);
    const b = table(["id", "name"], [["1", "Ada"], ["2", "Grace"], ["3", "Alan"]]);
    const result = diffSheets(a, b, { positional: true });
    expect(result.mode).toBe("positional");
    expect(result.added).toEqual([{ key: null, record: 4, values: ["3", "Alan"] }]);
    expect(result.warnings.map((warning) => warning.id)).toContain("positional");
  });

  it("never applies the positional fallback implicitly", () => {
    const result = diffSheets(A, A, {});
    expect(result.mode).toBe("none");
    expect(result.warnings.map((warning) => warning.id)).toContain("no-key");
    expect(result.changedCells).toHaveLength(0);
  });
});

describe("diffSheets — column changes", () => {
  it("matches headers case-insensitively and trimmed", () => {
    const a = table(["ID", " Name "], [["1", "Ada"]]);
    const b = table(["id", "name"], [["1", "Ada"]]);
    const result = diffSheets(a, b, { keyColumn: "ID" });
    expect(result.columns.pairs).toHaveLength(2);
    expect(result.columns.added).toHaveLength(0);
    expect(result.columns.removed).toHaveLength(0);
    expect(result.changedCells).toHaveLength(0);
  });

  it("reports added, removed and reordered columns", () => {
    const a = table(["id", "name", "plan"], [["1", "Ada", "pro"]]);
    const b = table(["plan", "id", "name", "region"], [["pro", "1", "Ada", "UK"]]);
    const result = diffSheets(a, b, { keyColumn: "id" });
    expect(result.columns.added.map((column) => column.bHeader)).toEqual(["region"]);
    expect(result.columns.removed).toHaveLength(0);
    expect(result.columns.reordered.map((entry) => entry.aHeader)).toEqual(["id", "name", "plan"]);
    // Reordering alone must not produce phantom cell changes.
    expect(result.changedCells).toHaveLength(0);
  });

  it("offers a labeled suggestion for an unmatched column but never applies it", () => {
    const a = table(["id", "First Name"], [["1", "Ada"]]);
    const b = table(["id", "first_name"], [["1", "Ada"]]);
    const result = diffSheets(a, b, { keyColumn: "id" });
    expect(result.columns.suggested[0]).toMatchObject({
      bHeader: "first_name",
      aHeader: "First Name",
    });
    expect(result.columns.pairs).toHaveLength(1);
    expect(result.changedCells).toHaveLength(0);
    expect(result.columns.suggested[0].reason.length).toBeGreaterThan(0);
  });

  it("applies an explicit manual mapping", () => {
    const a = table(["id", "First Name"], [["1", "Ada"]]);
    const b = table(["id", "first_name"], [["1", "Grace"]]);
    const result = diffSheets(a, b, {
      keyColumn: "id",
      columnMapping: { first_name: "First Name" },
    });
    expect(result.columns.pairs).toHaveLength(2);
    expect(result.changedCells).toEqual([
      {
        key: "1",
        aRecord: 2,
        bRecord: 2,
        column: "First Name",
        aColumn: 2,
        bColumn: 2,
        oldValue: "Ada",
        newValue: "Grace",
      },
    ]);
  });

  it("excludes duplicated header names and warns about them", () => {
    const a = table(["id", "name", "name"], [["1", "Ada", "Lovelace"]]);
    const b = table(["id", "name", "name"], [["1", "Ada", "Lovelace"]]);
    const result = diffSheets(a, b, { keyColumn: "id" });
    expect(result.columns.ambiguous).toEqual([{ name: "name", aIndexes: [2, 3], bIndexes: [2, 3] }]);
    expect(result.columns.pairs.map((pair) => pair.aHeader)).toEqual(["id"]);
    expect(result.warnings.map((warning) => warning.id)).toContain("duplicate-header-name");
  });

  it("warns when no column names match at all", () => {
    const a = table(["id", "name"], [["1", "Ada"]]);
    const b = table(["key", "label"], [["1", "Ada"]]);
    const result = diffSheets(a, b, {});
    expect(result.warnings.map((warning) => warning.id)).toContain("no-matching-columns");
  });

  it("diffColumns reports empty headers via the diff warnings", () => {
    const a = table(["id", ""], [["1", "x"]]);
    const b = table(["id", ""], [["1", "x"]]);
    const result = diffSheets(a, b, { keyColumn: "id" });
    expect(result.warnings.map((warning) => warning.id)).toContain("empty-header-File A");
  });
});

describe("comparison toggles", () => {
  it("defaults to raw, exact comparison", () => {
    expect(DEFAULT_TOGGLES).toEqual({
      trimWhitespace: false,
      ignoreCase: false,
      emptyEqualsMissing: false,
      numericAware: false,
    });
    expect(valuesDiffer(" Ada ", "Ada", DEFAULT_TOGGLES)).toBe(true);
    expect(valuesDiffer("Ada", "ada", DEFAULT_TOGGLES)).toBe(true);
    expect(valuesDiffer("1.0", "1", DEFAULT_TOGGLES)).toBe(true);
  });

  it("applies each toggle independently", () => {
    expect(valuesDiffer(" Ada ", "Ada", { ...DEFAULT_TOGGLES, trimWhitespace: true })).toBe(false);
    expect(valuesDiffer("Ada", "ada", { ...DEFAULT_TOGGLES, ignoreCase: true })).toBe(false);
    expect(valuesDiffer("1.0", "1", { ...DEFAULT_TOGGLES, numericAware: true })).toBe(false);
    expect(valuesDiffer("1.0", "1", { ...DEFAULT_TOGGLES, numericAware: false })).toBe(true);
    expect(valuesDiffer("+49", "49.0", { ...DEFAULT_TOGGLES, numericAware: true })).toBe(false);
    expect(valuesDiffer("1e2", "100", { ...DEFAULT_TOGGLES, numericAware: true })).toBe(false);
  });

  it("treats an empty value as a missing cell only when asked", () => {
    expect(valuesDiffer("", null, DEFAULT_TOGGLES)).toBe(true);
    expect(valuesDiffer("", null, { ...DEFAULT_TOGGLES, emptyEqualsMissing: true })).toBe(false);
    expect(valuesDiffer(null, null, DEFAULT_TOGGLES)).toBe(false);
  });

  it("records the enabled toggles on the result", () => {
    const result = diffSheets(A, A, {
      keyColumn: "id",
      toggles: { ignoreCase: true, numericAware: true },
    });
    expect(result.toggles).toEqual({
      trimWhitespace: false,
      ignoreCase: true,
      emptyEqualsMissing: false,
      numericAware: true,
    });
  });

  it("normalizes keys with the same toggles as cells", () => {
    const result = diffSheets(
      table(["id", "name"], [[" 1 ", "Ada"]]),
      table(["id", "name"], [["1", "Ada"]]),
      { keyColumn: "id", toggles: { trimWhitespace: true } },
    );
    expect(result.summary.matchedRows).toBe(1);
    expect(result.changedCells).toHaveLength(0);
    expect(normalizeKey(" 1 ", { ...DEFAULT_TOGGLES, trimWhitespace: true })).toBe("1");
    expect(normalizeKey("01", { ...DEFAULT_TOGGLES, numericAware: true })).toBe("1");
  });
});

describe("key suggestions", () => {
  it("suggests an identifier column whose values are unique in both files", () => {
    const suggestion = suggestKeyColumn(A, A, diffColumns(A, A).pairs);
    expect(suggestion?.header).toBe("id");
    expect(suggestion?.reason).toContain("unique");
  });

  it("prefers a column with unique values over one whose name merely looks like an id", () => {
    const a = table(["code", "name"], [["X", "Ada"], ["X", "Grace"]]);
    const suggestion = suggestKeyColumn(a, a, diffColumns(a, a).pairs);
    expect(suggestion?.header).toBe("name");
    expect(suggestion?.reason).toContain("unique");
  });

  it("still offers a suggestion when values are not unique, and labels it a guess", () => {
    const a = table(["code"], [["X"], ["X"]]);
    const suggestion = suggestKeyColumn(a, a, diffColumns(a, a).pairs);
    expect(suggestion?.header).toBe("code");
    expect(suggestion?.reason).toContain("not unique");
    expect(suggestion?.reason).toContain("duplicates will be reported");
  });

  it("returns null when no column is unique or looks like a key", () => {
    const a = table(["notes", "comments"], [["a", "a"], ["a", "a"]]);
    const suggestion = suggestKeyColumn(a, a, diffColumns(a, a).pairs);
    expect(suggestion).toBeNull();
  });

  it("names the suggested key in the no-key warning without applying it", () => {
    const result = diffSheets(A, A, {});
    const finding = result.warnings.find((warning) => warning.id === "no-key");
    expect(finding?.details[0]).toContain('Suggested key: "id"');
    expect(finding?.details[0]).toContain("not applied");
    expect(result.key.column).toBeNull();
  });
});
