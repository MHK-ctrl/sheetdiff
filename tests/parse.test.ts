import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { diffSheets } from "../src/diff";
import {
  detectFormat,
  fileSizeError,
  MAX_FILE_BYTES,
  parseCsvText,
  parseXlsxBuffer,
} from "../src/parse";
import { buildMarkdownReport } from "../src/report";

const here = dirname(fileURLToPath(import.meta.url));

describe("parseCsvText", () => {
  it("reads a clean CSV into headers and rows", () => {
    const parsed = parseCsvText("name,age\nBob,42\nAlice,37\n");
    expect(parsed.table.headers).toEqual(["name", "age"]);
    expect(parsed.table.rows).toEqual([
      ["Bob", "42"],
      ["Alice", "37"],
    ]);
    expect(parsed.stats.format).toBe("csv");
    expect(parsed.stats.delimiter).toBe(",");
    expect(parsed.stats.delimiterSource).toBe("auto");
    expect(parsed.findings.filter((finding) => finding.severity !== "info")).toEqual([]);
  });

  it("keeps quoted commas and newlines inside a single field", () => {
    const parsed = parseCsvText('name,note\n"Doe, Jane","line one\nline two"\n');
    expect(parsed.table.headers).toEqual(["name", "note"]);
    expect(parsed.table.rows[0]).toEqual(["Doe, Jane", "line one\nline two"]);
  });

  it("detects semicolon, tab and pipe files", () => {
    expect(parseCsvText("a;b\n1;2\n").stats.delimiter).toBe(";");
    expect(parseCsvText("a\tb\n1\t2\n").stats.delimiter).toBe("\t");
    expect(parseCsvText("a|b\n1|2\n").stats.delimiter).toBe("|");
  });

  it("labels a manual delimiter as manual", () => {
    const parsed = parseCsvText("a;b\n1;2\n", { delimiter: ";" });
    expect(parsed.stats.delimiterSource).toBe("manual");
    expect(parsed.stats.delimiterLabel).toBe("semicolon");
    expect(parsed.table.headers).toEqual(["a", "b"]);
  });

  it("warns when only one column is detected but another delimiter appears", () => {
    const parsed = parseCsvText("name;age\nBob;42\n", { delimiter: "," });
    expect(parsed.table.headers).toEqual(["name;age"]);
    const finding = parsed.findings.find((entry) => entry.id === "delimiter-guess");
    expect(finding?.severity).toBe("warning");
    expect(finding?.details.join(" ")).toContain("semicolon");
  });

  it("detects and reports a UTF-8 BOM without polluting the first header", () => {
    const parsed = parseCsvText("\uFEFFname,age\nBob,42\n");
    expect(parsed.stats.bomDetected).toBe(true);
    expect(parsed.table.headers).toEqual(["name", "age"]);
    expect(parsed.findings.find((finding) => finding.id === "bom")?.severity).toBe("info");
  });

  it("counts blank lines and skips them instead of treating them as rows", () => {
    const parsed = parseCsvText("a,b\n\n1,2\n\n");
    expect(parsed.stats.blankLineCount).toBe(2);
    expect(parsed.table.rows).toEqual([["1", "2"]]);
  });

  it("reports ragged rows with parsed record numbers, not editor line numbers", () => {
    const parsed = parseCsvText('a,b,c\n1,2,3\n4,5\n"multi\nline",7,8\n');
    expect(parsed.stats.raggedRowCount).toBe(1);
    const finding = parsed.findings.find((entry) => entry.id === "ragged");
    expect(finding?.details.join(" ")).toContain("Record 3: 2 fields (expected 3).");
    // The quoted newline stays inside one record, so the rows keep parsed-record order.
    expect(parsed.table.rows[1]).toEqual(["4", "5"]);
    expect(parsed.table.rows[2]).toEqual(["multi\nline", "7", "8"]);
  });

  it("explains unclosed quotes in plain language", () => {
    const parsed = parseCsvText('a,b\n1,"broken\n2,3\n');
    const finding = parsed.findings.find((entry) => entry.id === "parse-errors");
    expect(finding?.severity).toBe("error");
    const text = finding?.details.join(" ").toLowerCase() ?? "";
    expect(text).toContain("quote");
    expect(text).toContain("record 2");
  });

  it("notes mixed line endings", () => {
    const parsed = parseCsvText("a,b\r\n1,2\n3,4\r\n");
    expect(parsed.stats.lineEndings).toBe("mixed");
    expect(parsed.findings.find((finding) => finding.id === "mixed-line-endings")).toBeDefined();
  });

  it("handles empty and whitespace-only input without crashing", () => {
    const empty = parseCsvText("");
    expect(empty.table.headers).toEqual([]);
    expect(empty.table.rows).toEqual([]);
    expect(empty.findings.find((finding) => finding.id === "empty")?.severity).toBe("warning");
    expect(parseCsvText("   \n").findings.find((finding) => finding.id === "empty")).toBeDefined();
  });

  it("rejects files over the 5 MB limit via the size helper", () => {
    expect(fileSizeError(MAX_FILE_BYTES)).toBeNull();
    expect(fileSizeError(MAX_FILE_BYTES + 1)).toContain("5 MB");
  });

  it("picks the format from the file extension", () => {
    expect(detectFormat("data.csv")).toBe("csv");
    expect(detectFormat("DATA.XLSX")).toBe("xlsx");
    expect(detectFormat("pasted text")).toBe("csv");
  });
});

describe("parseXlsxBuffer", () => {
  const fixturePath = join(here, "fixtures", "people.xlsx");
  const buffer = readFileSync(fixturePath);
  const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);

  it("reads the first sheet, including inline strings and numbers", async () => {
    const parsed = await parseXlsxBuffer(arrayBuffer, { fileName: "people.xlsx" });
    expect(parsed.stats.format).toBe("xlsx");
    expect(parsed.stats.sheetNames).toEqual(["People", "Notes"]);
    expect(parsed.table.headers).toEqual(["id", "name", "plan", "mrr", "notes"]);
    expect(parsed.table.rows).toEqual([
      ["1001", "Ada Lovelace", "pro", "49", "Founding customer"],
      ["1002", "Grace Hopper", "pro", "49", ""],
      ["1003", "Alan Turing", "basic", "98", ""],
    ]);
  });

  it("does not evaluate formulas — it reads the last cached value", async () => {
    const parsed = await parseXlsxBuffer(arrayBuffer);
    // D4 is `<f>D2*2</f><v>98</v>`: the cached 98 is read, the formula is not run.
    expect(parsed.table.rows[2][3]).toBe("98");
    expect(parsed.findings.some((finding) => finding.details.join(" ").includes("formulas are not evaluated"))).toBe(true);
  });

  it("reads a named sheet and reports multi-sheet workbooks", async () => {
    const parsed = await parseXlsxBuffer(arrayBuffer, { sheet: "Notes" });
    expect(parsed.stats.sheetName).toBe("Notes");
    expect(parsed.table.headers).toEqual(["Sheet two heading"]);
    expect(parsed.findings[0].details.join(" ")).toContain("2 sheets");
  });

  it("reports the first sheet by name when no sheet is chosen", async () => {
    const parsed = await parseXlsxBuffer(arrayBuffer);
    expect(parsed.stats.sheetName).toBe("People");
  });

  it("diffs two xlsx sheets end to end", async () => {
    const a = await parseXlsxBuffer(arrayBuffer);
    const b = await parseXlsxBuffer(arrayBuffer);
    const result = diffSheets(a.table, b.table, { keyColumn: "id" });
    expect(result.summary.matchedRows).toBe(3);
    expect(result.changedCells).toHaveLength(0);
  });
});

describe("buildMarkdownReport", () => {
  it("includes the summary, both raw values, toggles and the record-number caveat", () => {
    const a = parseCsvText("id,plan\n1,pro\n2,basic\n", { fileName: "a.csv" });
    const b = parseCsvText("id,plan\n1,team\n3,basic\n", { fileName: "b.csv" });
    const result = diffSheets(a.table, b.table, {
      keyColumn: "id",
      toggles: { trimWhitespace: true },
    });
    const markdown = buildMarkdownReport(result, a, b, {
      aName: "a.csv",
      bName: "b.csv",
      generatedAt: "2026-10-03T00:00:00.000Z",
      repoUrl: "https://github.com/MHK-ctrl/sheetdiff",
    });
    expect(markdown).toContain("# SheetDiff report");
    expect(markdown).toContain("File A (original):** a.csv");
    expect(markdown).toContain("File B (updated):** b.csv");
    expect(markdown).toContain('by key column “id”');
    expect(markdown).toContain("trim whitespace");
    expect(markdown).toContain("| 1 | plan | pro | team | 2 | 2 |");
    expect(markdown).toContain("parsed record numbers");
    expect(markdown).toContain("Generated with [SheetDiff](https://github.com/MHK-ctrl/sheetdiff)");
  });

  it("says explicitly when no toggles were enabled", () => {
    const a = parseCsvText("id\n1\n", { fileName: "a.csv" });
    const b = parseCsvText("id\n1\n", { fileName: "b.csv" });
    const result = diffSheets(a.table, b.table, { keyColumn: "id" });
    const markdown = buildMarkdownReport(result, a, b, {
      aName: "a.csv",
      bName: "b.csv",
      generatedAt: "2026-10-03T00:00:00.000Z",
    });
    expect(markdown).toContain("none — raw values compared exactly");
  });

  it("escapes pipes and newlines in values so the table stays valid", () => {
    const a = parseCsvText("id,note\n1,hello\n", { fileName: "a.csv" });
    const b = parseCsvText("id,note\n1,pipe|here\n", { fileName: "b.csv" });
    const result = diffSheets(a.table, b.table, { keyColumn: "id" });
    const markdown = buildMarkdownReport(result, a, b, {
      aName: "a.csv",
      bName: "b.csv",
      generatedAt: "2026-10-03T00:00:00.000Z",
    });
    expect(markdown).toContain("pipe\\|here");
  });
});
