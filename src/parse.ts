import Papa from "papaparse";
import readXlsxFile, { readSheetNames } from "read-excel-file";
import type { Finding, SheetTable } from "./diff";

/** Files larger than this are rejected before any parsing happens. */
export const MAX_FILE_BYTES = 5 * 1024 * 1024; // 5 MB

export type SheetFormat = "csv" | "xlsx";
export type LineEnding = "lf" | "crlf" | "cr" | "mixed" | "none";

export interface SheetStats {
  fileName: string;
  format: SheetFormat;
  /** Size of the input in UTF-8 bytes (or the file size for .xlsx). */
  byteSize: number;
  /** CSV only: the delimiter that was used. */
  delimiter: string;
  delimiterSource: "auto" | "manual" | "n/a";
  delimiterLabel: string;
  bomDetected: boolean;
  lineEndings: LineEnding;
  /** CSV only: records that were blank lines. */
  blankLineCount: number;
  /** CSV only: records whose field count differs from the header. */
  raggedRowCount: number;
  /** CSV only: Papa Parse errors (excluding "undetectable delimiter"). */
  parseErrorCount: number;
  /** XLSX only: the sheet that was read. */
  sheetName: string | null;
  /** XLSX only: all sheet names in the workbook. */
  sheetNames: string[];
}

export interface ParsedSheet {
  table: SheetTable;
  stats: SheetStats;
  /** Things the user should know about how the file was read. */
  findings: Finding[];
}

export interface CsvOptions {
  /** Force a delimiter. Omit for auto-detection (which is always a guess). */
  delimiter?: string;
  fileName?: string;
}

export interface XlsxOptions {
  /** Sheet name or 0-based index. Defaults to the first sheet. */
  sheet?: string | number;
  fileName?: string;
}

const DELIMITER_NAMES: Record<string, string> = {
  ",": "comma",
  ";": "semicolon",
  "\t": "tab",
  "|": "pipe",
};

const CANDIDATE_DELIMITERS = [",", ";", "\t", "|"];

const PARSE_ERROR_HINTS: Record<string, string> = {
  MissingQuotes:
    'A quoted field is never closed — look for a missing double quote ("), usually on this record or the one before it.',
  InvalidQuotes:
    'A quote appears in the middle of a field. Wrap the field in quotes and escape inner quotes by doubling them ("").',
};

export function delimiterLabel(delimiter: string): string {
  return DELIMITER_NAMES[delimiter] ?? JSON.stringify(delimiter);
}

/** Returns a human-readable error when `bytes` exceeds the limit, else null. */
export function fileSizeError(bytes: number): string | null {
  if (bytes <= MAX_FILE_BYTES) return null;
  const mb = (bytes / (1024 * 1024)).toFixed(1);
  return (
    `This file is ${mb} MB, and SheetDiff only reads files up to 5 MB because everything ` +
    "runs in your browser tab. Nothing was uploaded — try splitting the file first."
  );
}

/** Decide how to read a file from its name. Defaults to CSV for anything unknown. */
export function detectFormat(fileName: string): SheetFormat {
  return fileName.toLowerCase().endsWith(".xlsx") ? "xlsx" : "csv";
}

export function utf8ByteLength(text: string): number {
  return new TextEncoder().encode(text).length;
}

function detectLineEndings(text: string): LineEnding {
  const crlf = (text.match(/\r\n/g) ?? []).length;
  const lf = (text.match(/\n/g) ?? []).length - crlf;
  const cr = (text.match(/\r/g) ?? []).length - crlf;
  const kinds: LineEnding[] = [];
  if (crlf > 0) kinds.push("crlf");
  if (lf > 0) kinds.push("lf");
  if (cr > 0) kinds.push("cr");
  if (kinds.length === 0) return "none";
  return kinds.length === 1 ? kinds[0] : "mixed";
}

function isBlankRecord(record: string[]): boolean {
  return record.length === 0 || (record.length === 1 && record[0] === "");
}

/**
 * Read a CSV (or TSV / semicolon / pipe) text into a sheet.
 *
 * Pure with respect to the input: the text is never modified. Record numbers in
 * the diff are *parsed record* numbers (1-based, header = record 1), not editor
 * line numbers — a newline inside a quoted field shifts the numbering.
 */
export function parseCsvText(input: string, options: CsvOptions = {}): ParsedSheet {
  const fileName = options.fileName ?? "pasted text";
  const manualDelimiter =
    options.delimiter && options.delimiter.length > 0 ? options.delimiter : undefined;
  const bomDetected = input.charCodeAt(0) === 0xfeff;
  const text = bomDetected ? input.slice(1) : input;
  const byteSize = utf8ByteLength(input);
  const lineEndings = detectLineEndings(text);
  const findings: Finding[] = [];

  const baseStats = (delimiter: string, source: SheetStats["delimiterSource"]): SheetStats => ({
    fileName,
    format: "csv",
    byteSize,
    delimiter,
    delimiterSource: source,
    delimiterLabel: delimiterLabel(delimiter),
    bomDetected,
    lineEndings,
    blankLineCount: 0,
    raggedRowCount: 0,
    parseErrorCount: 0,
    sheetName: null,
    sheetNames: [],
  });

  if (bomDetected) {
    findings.push({
      id: "bom",
      severity: "info",
      title: "UTF-8 BOM detected",
      details: [
        "The file starts with a byte order mark (EF BB BF). It is ignored while reading — the bytes on disk are not touched — but some tools include it in the first column name.",
      ],
    });
  }

  if (text.trim() === "") {
    findings.push({
      id: "empty",
      severity: "warning",
      title: "This file is empty",
      details: ["There is nothing to compare. If you expected data, the file or paste may have been empty."],
    });
    return {
      table: { headers: [], rows: [] },
      stats: baseStats(manualDelimiter ?? ",", manualDelimiter ? "manual" : "auto"),
      findings,
    };
  }

  // Delimiter handling: auto-detect, but on a pass that ignores blank rows —
  // otherwise a trailing newline drags the guesser's average field count down
  // and detection fails for ordinary comma/semicolon/tab/pipe files. Detection
  // is still only a guess, and the UI always offers a manual override.
  let delimiter = manualDelimiter ?? "";
  if (delimiter === "") {
    const detection = Papa.parse<string[]>(text, { skipEmptyLines: true, preview: 10 });
    delimiter = detection.meta.delimiter || ",";
  }

  const parsed = Papa.parse<string[]>(text, {
    delimiter,
    skipEmptyLines: false,
    header: false,
    dynamicTyping: false,
  });

  const records = parsed.data.slice();
  // A file that ends with a line terminator parses to a final empty record;
  // that is the normal end of the file, not a blank line.
  if (text.endsWith("\n") || text.endsWith("\r")) {
    const last = records[records.length - 1];
    if (last && isBlankRecord(last)) records.pop();
  }

  const firstNonBlank = records.findIndex((record) => !isBlankRecord(record));
  const blankLineCount = records.filter(isBlankRecord).length;

  if (firstNonBlank === -1) {
    return {
      table: { headers: [], rows: [] },
      stats: baseStats(delimiter, manualDelimiter ? "manual" : "auto"),
      findings,
    };
  }

  const headerRecord = records[firstNonBlank];
  const dataRecords = records.slice(firstNonBlank + 1);
  const columnCount = headerRecord.length;

  if (columnCount === 1) {
    const firstLine = text.split(/\r\n|\r|\n/, 1)[0] ?? "";
    const alternatives = CANDIDATE_DELIMITERS.filter(
      (candidate) => candidate !== delimiter && firstLine.includes(candidate),
    );
    if (alternatives.length > 0) {
      const names = alternatives.map((candidate) => `a ${delimiterLabel(candidate)}`).join(" or ");
      findings.push({
        id: "delimiter-guess",
        severity: "warning",
        title: "Only one column detected — the delimiter guess looks wrong",
        details: [
          `The first line contains ${names}, but the file was read as a single column.`,
          `If ${names} is the real delimiter, choose it in the delimiter menu for this file and the diff will re-run.`,
        ],
      });
    }
  }

  const parseErrors = parsed.errors.filter((error) => error.code !== "UndetectableDelimiter");
  if (parseErrors.length > 0) {
    const details = parseErrors.slice(0, 20).map((error) => {
      const where = typeof error.row === "number" ? `Record ${error.row + 1}: ` : "";
      const hint = PARSE_ERROR_HINTS[error.code];
      return `${where}${error.message}${hint ? ` — ${hint}` : ""}`;
    });
    if (parseErrors.length > 20) {
      details.push(`…and ${parseErrors.length - 20} more parser errors.`);
    }
    findings.push({
      id: "parse-errors",
      severity: "error",
      title: `Parser errors: ${parseErrors.length}`,
      details,
    });
  }

  // Ragged rows: compared against the header record. Blank lines are counted
  // separately and never treated as ragged.
  const raggedRows: { record: number; expected: number; actual: number }[] = [];
  dataRecords.forEach((record, index) => {
    if (isBlankRecord(record)) return;
    if (record.length !== columnCount) {
      raggedRows.push({ record: firstNonBlank + 1 + index + 1, expected: columnCount, actual: record.length });
    }
  });
  if (raggedRows.length > 0) {
    const shown = raggedRows.slice(0, 20);
    const details = shown.map(
      (row) => `Record ${row.record}: ${row.actual} field${row.actual === 1 ? "" : "s"} (expected ${row.expected}).`,
    );
    if (raggedRows.length > shown.length) {
      details.push(`…and ${raggedRows.length - shown.length} more.`);
    }
    details.unshift(
      "Missing cells are shown as “no value” and are compared as such — nothing was added or removed.",
    );
    findings.push({
      id: "ragged",
      severity: "warning",
      title: `Ragged rows: ${raggedRows.length} record${raggedRows.length === 1 ? "" : "s"} with a different field count`,
      details,
    });
  }

  if (blankLineCount > 0) {
    findings.push({
      id: "blank-lines",
      severity: "info",
      title: `Blank lines: ${blankLineCount}`,
      details: ["Blank lines are skipped, not treated as ragged rows."],
    });
  }

  if (lineEndings === "mixed") {
    findings.push({
      id: "mixed-line-endings",
      severity: "info",
      title: "Mixed line endings",
      details: [
        "The file contains both Windows-style (CRLF) and Unix-style (LF) line endings. It was read as-is; most tools cope with this.",
      ],
    });
  }

  if (text.includes("\uFFFD")) {
    findings.push({
      id: "replacement-characters",
      severity: "info",
      title: "Characters that are not valid UTF-8",
      details: [
        'The text contains "\uFFFD" replacement characters. SheetDiff reads UTF-8 only — if this file uses another encoding, not all bytes could be read.',
      ],
    });
  }

  const stats: SheetStats = {
    ...baseStats(delimiter, manualDelimiter ? "manual" : "auto"),
    blankLineCount,
    raggedRowCount: raggedRows.length,
    parseErrorCount: parseErrors.length,
  };

  return {
    table: {
      headers: headerRecord.slice(),
      rows: dataRecords.filter((record) => !isBlankRecord(record)).map((record) => record.slice()),
    },
    stats,
    findings,
  };
}

function cellToText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value);
}

/**
 * Read the sheet names of an .xlsx workbook, returning an empty list if the
 * workbook cannot be read. Sheet names are only used to label the sheet picker
 * and the report, so a failure here must not fail the whole parse.
 */
export async function readXlsxSheetNames(buffer: ArrayBuffer): Promise<string[]> {
  return readSheetNames(buffer).catch(() => []);
}

/**
 * Read an .xlsx workbook into a sheet. Only stored (cached) cell values are
 * read: formulas are NOT evaluated, and a formula cell shows its last cached
 * result. Trailing empty rows and columns are dropped by the reader, so record
 * numbers are parsed-record numbers.
 */
export async function parseXlsxBuffer(
  buffer: ArrayBuffer,
  options: XlsxOptions = {},
): Promise<ParsedSheet> {
  const fileName = options.fileName ?? "workbook.xlsx";
  const byteSize = buffer.byteLength;
  const findings: Finding[] = [];
  const sheetNames = await readXlsxSheetNames(buffer);

  const rows = await readXlsxFile(buffer, { sheet: options.sheet ?? 1, trim: false });
  const asText = rows.map((row) => row.map(cellToText));

  findings.push({
    id: "xlsx-read",
    severity: "info",
    title: "Excel workbook read",
    details: [
      "Only stored cell values are read — formulas are not evaluated, and a formula cell shows its last saved result.",
      ...(sheetNames.length > 1
        ? [`The workbook has ${sheetNames.length} sheets (${sheetNames.join(", ")}); only one sheet is compared at a time.`]
        : []),
    ],
  });

  const firstNonEmpty = asText.findIndex((row) => row.some((cell) => cell.trim() !== ""));
  const headerRecord = firstNonEmpty === -1 ? [] : asText[firstNonEmpty];
  const dataRows = firstNonEmpty === -1 ? [] : asText.slice(firstNonEmpty + 1);
  const keptRows = dataRows.filter((row) => row.some((cell) => cell.trim() !== ""));
  if (dataRows.length !== keptRows.length) {
    findings.push({
      id: "xlsx-empty-rows",
      severity: "info",
      title: `Empty rows skipped: ${dataRows.length - keptRows.length}`,
      details: ["Rows where every cell is empty are skipped — they are not treated as removed or added rows."],
    });
  }

  const sheetName =
    typeof options.sheet === "string"
      ? options.sheet
      : sheetNames[options.sheet === undefined ? 0 : Number(options.sheet) - 1] ?? null;

  return {
    table: {
      headers: headerRecord.slice(),
      rows: keptRows.map((row) => row.slice()),
    },
    stats: {
      fileName,
      format: "xlsx",
      byteSize,
      delimiter: "",
      delimiterSource: "n/a",
      delimiterLabel: "n/a",
      bomDetected: false,
      lineEndings: "none",
      blankLineCount: 0,
      raggedRowCount: 0,
      parseErrorCount: 0,
      sheetName,
      sheetNames,
    },
    findings,
  };
}
