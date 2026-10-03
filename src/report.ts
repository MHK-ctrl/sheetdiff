import type { CellChange, DiffResult, Finding, RowChange } from "./diff";
import type { ParsedSheet } from "./parse";

export interface ReportMeta {
  /** File name of file A. */
  aName: string;
  /** File name of file B. */
  bName: string;
  /** Human-readable timestamp shown in the report header. */
  generatedAt: string;
  /** Public repo URL, linked in the footer. */
  repoUrl?: string;
}

/** Detail rows per category before the report summarizes the rest. */
export const MAX_REPORT_DETAIL_ROWS = 200;

export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} bytes`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/** Markdown table cells must not contain pipes or newlines. */
function cell(value: string | null): string {
  if (value === null || value === "") return "—";
  return value.replace(/\|/g, "\\|").replace(/\r?\n/g, " ");
}

function findingLines(findings: Finding[]): string[] {
  const lines: string[] = [];
  for (const finding of findings) {
    lines.push(`### ${finding.title}`, "");
    for (const detail of finding.details) lines.push(`- ${detail}`);
    lines.push("");
  }
  return lines;
}

function describeSheet(sheet: ParsedSheet): string {
  const stats = sheet.stats;
  if (stats.format === "xlsx") {
    return `Excel workbook${stats.sheetName ? `, sheet “${stats.sheetName}”` : ""}`;
  }
  return `${stats.delimiterLabel} (${stats.delimiterSource === "manual" ? "chosen manually" : "auto-detected guess"})`;
}

function describeToggles(result: DiffResult): string[] {
  const { toggles } = result;
  const enabled: string[] = [];
  if (toggles.trimWhitespace) enabled.push("trim whitespace");
  if (toggles.ignoreCase) enabled.push("case-insensitive");
  if (toggles.emptyEqualsMissing) enabled.push("empty value = missing cell");
  if (toggles.numericAware) enabled.push("numeric-aware (1.0 = 1)");
  return enabled;
}

function rowChangeLines(changes: RowChange[], headers: string[]): string[] {
  if (changes.length === 0) return ["None.", ""];
  const shown = changes.slice(0, MAX_REPORT_DETAIL_ROWS);
  const lines = [
    `| Record | ${headers.map(cell).join(" | ")} |`,
    `| --- | ${headers.map(() => "---").join(" | ")} |`,
  ];
  for (const change of shown) {
    const key = change.key === null ? "—" : cell(change.key);
    const values = headers.map((_, index) => cell(change.values[index] ?? null));
    lines.push(`| ${change.record} (key ${key}) | ${values.join(" | ")} |`);
  }
  lines.push("");
  if (changes.length > shown.length) {
    lines.push(`…and ${changes.length - shown.length} more row${changes.length - shown.length === 1 ? "" : "s"}.`, "");
  }
  return lines;
}

function changedCellLines(changes: CellChange[]): string[] {
  if (changes.length === 0) return ["None.", ""];
  const shown = changes.slice(0, MAX_REPORT_DETAIL_ROWS);
  const lines = [
    "| Key | Column | Old value | New value | A record | B record |",
    "| --- | --- | --- | --- | --- | --- |",
  ];
  for (const change of shown) {
    lines.push(
      `| ${cell(change.key)} | ${cell(change.column)} | ${cell(change.oldValue)} | ${cell(change.newValue)} | ${change.aRecord} | ${change.bRecord} |`,
    );
  }
  lines.push("");
  if (changes.length > shown.length) {
    lines.push(`…and ${changes.length - shown.length} more changed cell${changes.length - shown.length === 1 ? "" : "s"}.`, "");
  }
  return lines;
}

/**
 * Render a diff as a Markdown report suitable for pasting into a PR or an
 * issue. Both raw values are always shown; the enabled comparison toggles are
 * recorded so the reader knows exactly what was compared.
 */
export function buildMarkdownReport(
  result: DiffResult,
  a: ParsedSheet,
  b: ParsedSheet,
  meta: ReportMeta,
): string {
  const summary = result.summary;
  const lines: string[] = [
    "# SheetDiff report",
    "",
    `- **File A (original):** ${meta.aName} — ${describeSheet(a)} — ${formatBytes(a.stats.byteSize)}`,
    `- **File B (updated):** ${meta.bName} — ${describeSheet(b)} — ${formatBytes(b.stats.byteSize)}`,
    `- **Generated:** ${meta.generatedAt}`,
    `- **Row matching:** ${
      result.mode === "key"
        ? `by key column “${result.key.column}”`
        : result.mode === "positional"
          ? "by position (explicit fallback — can be wrong when rows were inserted or deleted)"
          : "not performed (no key column selected)"
    }`,
  ];
  const enabledToggles = describeToggles(result);
  lines.push(
    `- **Comparison toggles:** ${enabledToggles.length === 0 ? "none — raw values compared exactly" : enabledToggles.join(", ")}`,
  );
  lines.push("");
  lines.push(
    `**${summary.addedRows} added · ${summary.removedRows} removed · ${summary.changedRows} rows changed (${summary.changedCells} cells) · ${summary.matchedRows} rows matched**`,
  );
  lines.push("");
  lines.push(
    "> Record numbers are parsed record numbers (record 1 is the header row),",
    "> not editor line numbers — a newline inside a quoted field shifts the numbering.",
    "",
  );

  if (enabledToggles.length > 0) {
    lines.push(
      "> The comparison toggles above were enabled by the user. The raw values below are",
      "> shown exactly as they appear in the files; nothing was normalized in the files themselves.",
      "",
    );
  }

  const warnings = result.warnings;
  const notes = result.notes;
  const parseFindings = [...a.findings, ...b.findings];

  lines.push(`## Summary`, "");
  lines.push(`- Rows in file A: ${summary.aRows}`);
  lines.push(`- Rows in file B: ${summary.bRows}`);
  lines.push(`- Rows matched: ${summary.matchedRows}`);
  lines.push(`- Added rows (in B, not A): ${summary.addedRows}`);
  lines.push(`- Removed rows (in A, not B): ${summary.removedRows}`);
  lines.push(`- Rows with changed cells: ${summary.changedRows}`);
  lines.push(`- Changed cells: ${summary.changedCells}`);
  lines.push(`- Duplicate keys excluded: ${summary.ambiguousKeys}`);
  lines.push(`- Rows with an empty key excluded: ${summary.emptyKeyRowsA + summary.emptyKeyRowsB}`);
  lines.push(`- Columns matched: ${result.columns.pairs.length}`);
  lines.push("");

  lines.push(`## Added rows (${summary.addedRows})`, "");
  lines.push(...rowChangeLines(result.added, b.table.headers));
  lines.push(`## Removed rows (${summary.removedRows})`, "");
  lines.push(...rowChangeLines(result.removed, a.table.headers));
  lines.push(`## Changed cells (${summary.changedCells})`, "");
  lines.push(...changedCellLines(result.changedCells));

  lines.push(`## Column changes`, "");
  if (result.columns.added.length > 0) {
    lines.push(
      `- Added columns (in B): ${result.columns.added.map((column) => `“${column.bHeader}” (column ${column.bIndex + 1})`).join(", ")}`,
    );
  }
  if (result.columns.removed.length > 0) {
    lines.push(
      `- Removed columns (in A): ${result.columns.removed.map((column) => `“${column.aHeader}” (column ${column.aIndex + 1})`).join(", ")}`,
    );
  }
  if (result.columns.reordered.length > 0) {
    lines.push(
      `- Reordered columns: ${result.columns.reordered.map((entry) => `“${entry.aHeader}” is column ${entry.aColumn} in A and ${entry.bColumn} in B`).join("; ")}`,
    );
  }
  if (result.columns.suggested.length > 0) {
    lines.push(
      "- Suggested mappings (guesses, not applied): " +
        result.columns.suggested
          .map((entry) => `“${entry.bHeader}” (B) ↔ “${entry.aHeader}” (A) — ${entry.reason}`)
          .join("; "),
    );
  }
  if (result.columns.ambiguous.length > 0) {
    lines.push(
      "- Duplicate header names (excluded from comparison): " +
        result.columns.ambiguous.map((entry) => `“${entry.name}”`).join(", "),
    );
  }
  if (
    result.columns.added.length === 0 &&
    result.columns.removed.length === 0 &&
    result.columns.reordered.length === 0
  ) {
    lines.push("No column-level changes.");
  }
  lines.push("");

  if (warnings.length > 0) {
    lines.push(`## ⚠️ Warnings (${warnings.length})`, "");
    lines.push(...findingLines(warnings));
  }
  if (notes.length > 0) {
    lines.push(`## ℹ️ Notes (${notes.length})`, "");
    lines.push(...findingLines(notes));
  }
  if (parseFindings.length > 0) {
    lines.push(`## How the files were read`, "");
    lines.push(...findingLines(parseFindings));
  }

  lines.push("---");
  lines.push(
    meta.repoUrl
      ? `Generated with [SheetDiff](${meta.repoUrl}) — everything ran in the browser; neither file was uploaded.`
      : "Generated with SheetDiff — everything ran in the browser; neither file was uploaded.",
  );
  return `${lines.join("\n")}\n`;
}
