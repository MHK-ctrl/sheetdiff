/**
 * Pure diff engine for SheetDiff. No DOM, no network, no globals.
 *
 * Nothing here ever modifies the input tables: the engine reports differences,
 * it never "repairs" data. Every guess (suggested key column, suggested column
 * mapping, positional fallback) is labeled as a guess and must be confirmed by
 * the caller.
 */

export type Severity = "error" | "warning" | "info";

export interface Finding {
  id: string;
  severity: Severity;
  title: string;
  details: string[];
}

/** A parsed sheet: header names plus data rows (raw strings, never normalized). */
export interface SheetTable {
  /** Header names exactly as they appear in the file (not trimmed). */
  headers: string[];
  /**
   * Data rows. A row may be shorter than `headers`; a missing cell is
   * `undefined` and is displayed as "no value".
   */
  rows: string[][];
}

/**
 * Comparison toggles. All are OFF by default and are recorded in the report.
 * Toggle order in the pipeline is: trim -> case-fold -> numeric -> empty/missing.
 */
export interface CompareToggles {
  trimWhitespace: boolean;
  ignoreCase: boolean;
  emptyEqualsMissing: boolean;
  numericAware: boolean;
}

export const DEFAULT_TOGGLES: CompareToggles = {
  trimWhitespace: false,
  ignoreCase: false,
  emptyEqualsMissing: false,
  numericAware: false,
};

export interface ColumnPair {
  aHeader: string;
  aIndex: number;
  bHeader: string;
  bIndex: number;
}

export interface ColumnDiff {
  /** Columns matched between the two files (automatic + manual mapping). */
  pairs: ColumnPair[];
  /** Columns only in file B. */
  added: { bHeader: string; bIndex: number }[];
  /** Columns only in file A. */
  removed: { aHeader: string; aIndex: number }[];
  /** Paired columns whose relative order changed. */
  reordered: { aHeader: string; bHeader: string; aColumn: number; bColumn: number }[];
  /** Header names that appear more than once in a file (excluded from pairing). */
  ambiguous: { name: string; aIndexes: number[]; bIndexes: number[] }[];
  /** Guessed mappings for unmatched columns. Never applied automatically. */
  suggested: { bHeader: string; bIndex: number; aHeader: string; aIndex: number; reason: string }[];
}

export interface KeySuggestion {
  header: string;
  reason: string;
}

export interface KeyResolution {
  /** "key" = rows matched by key column, "positional" = explicit fallback, "none" = no comparison. */
  mode: "key" | "positional" | "none";
  /** Header name (in file A) used as the key, when mode is "key". */
  column: string | null;
  /** A guess offered to the user; never applied automatically. */
  suggested: KeySuggestion | null;
  /** Set when a requested key column could not be used. */
  problem: string | null;
}

export interface RowChange {
  /** Normalized key value (null in positional mode). */
  key: string | null;
  /** 1-based parsed record number in the file the row came from. */
  record: number;
  /** Cell values aligned to that file's headers; null = no value at that position. */
  values: (string | null)[];
}

export interface CellChange {
  key: string;
  /** Record number in file A. */
  aRecord: number;
  /** Record number in file B. */
  bRecord: number;
  /** Column name in file A (canonical name for the pair). */
  column: string;
  aColumn: number;
  bColumn: number;
  oldValue: string | null;
  newValue: string | null;
}

export interface AmbiguousKey {
  key: string;
  aRecords: number[];
  bRecords: number[];
  reason: "duplicate-in-a" | "duplicate-in-b" | "duplicate-in-both";
}

export interface DiffSummary {
  aRows: number;
  bRows: number;
  matchedRows: number;
  addedRows: number;
  removedRows: number;
  /** Matched rows with at least one changed cell. */
  changedRows: number;
  changedCells: number;
  ambiguousKeys: number;
  /** Rows that could not be matched because the key value is empty. */
  emptyKeyRowsA: number;
  emptyKeyRowsB: number;
}

export interface DiffResult {
  mode: "key" | "positional" | "none";
  toggles: CompareToggles;
  key: KeyResolution;
  columns: ColumnDiff;
  summary: DiffSummary;
  added: RowChange[];
  removed: RowChange[];
  changedCells: CellChange[];
  ambiguousKeys: AmbiguousKey[];
  emptyKey: { a: number[]; b: number[] };
  /** Matched rows whose position changed but whose content did not. */
  moved: { key: string; aRecord: number; bRecord: number }[];
  /** Things the user must fix or confirm (severity error/warning). */
  warnings: Finding[];
  /** Informational observations only. */
  notes: Finding[];
}

export interface DiffOptions {
  /** Header name from file A to use as the key. */
  keyColumn?: string | null;
  /** Manual mapping: file B header -> file A header. Applied before auto-matching. */
  columnMapping?: Record<string, string>;
  toggles?: Partial<CompareToggles>;
  /** Explicit fallback: compare rows by position when no usable key exists. */
  positional?: boolean;
}

/** Matches a strict decimal number. Surrounding whitespace is not accepted. */
const NUMERIC_RE = /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/;

/** Header names that usually identify a row. Used only to rank suggestions. */
const IDENTIFIER_RE = /^(id|.*_id|.*-id|.*id|uuid|guid|key|.*_key|sku|code|email|e-mail)$/i;

/** How many detail lines a finding may carry before it is summarized. */
const MAX_FINDING_DETAILS = 20;

/** 1-based parsed record number for a data-row index (record 1 is the header). */
export function dataRowRecord(rowIndex: number): number {
  return rowIndex + 2;
}

function describeRecords(records: number[]): string {
  const shown = records.slice(0, MAX_FINDING_DETAILS);
  const text = shown.join(", ");
  return records.length > shown.length ? `${text}, … (${records.length - shown.length} more)` : text;
}

export function normalizeHeader(header: string): string {
  return header.trim().toLowerCase();
}

/** Compare two cell values under the enabled toggles. */
export function valuesDiffer(
  a: string | null,
  b: string | null,
  toggles: CompareToggles,
): boolean {
  const stage = (value: string | null): string | undefined => {
    if (value === null) return undefined;
    let out = value;
    if (toggles.trimWhitespace) out = out.trim();
    if (toggles.ignoreCase) out = out.toLowerCase();
    return out;
  };
  let x = stage(a);
  let y = stage(b);
  if (toggles.emptyEqualsMissing) {
    if (x === undefined || x === "") x = "";
    if (y === undefined || y === "") y = "";
  }
  if (x === undefined || y === undefined) return x !== y;
  if (toggles.numericAware && NUMERIC_RE.test(x) && NUMERIC_RE.test(y)) {
    return Number(x) !== Number(y);
  }
  return x !== y;
}

/**
 * Normalize a key value for matching. Applies trim, case-folding and
 * numeric-aware normalization (the same toggles as cell comparison, except
 * empty-equals-missing, which only concerns missing cells).
 */
export function normalizeKey(value: string | null, toggles: CompareToggles): string {
  if (value === null) return "";
  let out = value;
  if (toggles.trimWhitespace) out = out.trim();
  if (toggles.ignoreCase) out = out.toLowerCase();
  if (toggles.numericAware && NUMERIC_RE.test(out)) out = String(Number(out));
  return out;
}

function cellAt(table: SheetTable, rowIndex: number, columnIndex: number): string | null {
  const row = table.rows[rowIndex];
  if (!row) return null;
  const value = row[columnIndex];
  return value === undefined ? null : value;
}

function cellsAt(table: SheetTable, rowIndex: number): (string | null)[] {
  return table.headers.map((_, columnIndex) => cellAt(table, rowIndex, columnIndex));
}

function headerIndexes(table: SheetTable): Map<string, number[]> {
  const map = new Map<string, number[]>();
  table.headers.forEach((header, index) => {
    const key = normalizeHeader(header);
    const existing = map.get(key);
    if (existing) existing.push(index);
    else map.set(key, [index]);
  });
  return map;
}

/**
 * Match columns between the two files. Header names are matched trimmed and
 * case-insensitively; anything else is reported as added, removed or ambiguous,
 * with a labeled guess offered for manual mapping.
 */
export function diffColumns(
  a: SheetTable,
  b: SheetTable,
  columnMapping: Record<string, string> = {},
): ColumnDiff {
  const aByName = headerIndexes(a);
  const bByName = headerIndexes(b);
  const ambiguousNames = new Set<string>();
  for (const [name, indexes] of aByName) if (indexes.length > 1) ambiguousNames.add(name);
  for (const [name, indexes] of bByName) if (indexes.length > 1) ambiguousNames.add(name);

  const pairs: ColumnPair[] = [];
  const usedA = new Set<number>();
  const usedB = new Set<number>();

  // 1. Manual mappings win, so a deliberate choice is never overridden.
  for (const [bHeader, aHeader] of Object.entries(columnMapping)) {
    const aIndex = a.headers.indexOf(aHeader);
    const bIndex = b.headers.indexOf(bHeader);
    if (aIndex === -1 || bIndex === -1) continue;
    if (usedA.has(aIndex) || usedB.has(bIndex)) continue;
    usedA.add(aIndex);
    usedB.add(bIndex);
    pairs.push({ aHeader: a.headers[aIndex], aIndex, bHeader: b.headers[bIndex], bIndex });
  }

  // 2. Automatic matching by normalized header name, skipping ambiguous names.
  a.headers.forEach((aHeader, aIndex) => {
    const name = normalizeHeader(aHeader);
    if (ambiguousNames.has(name) || usedA.has(aIndex)) return;
    const candidates = bByName.get(name) ?? [];
    const bIndex = candidates[0];
    if (bIndex === undefined || usedB.has(bIndex)) return;
    usedA.add(aIndex);
    usedB.add(bIndex);
    pairs.push({ aHeader, aIndex, bHeader: b.headers[bIndex], bIndex });
  });

  pairs.sort((left, right) => left.aIndex - right.aIndex);

  const added = b.headers
    .map((bHeader, bIndex) => ({ bHeader, bIndex }))
    .filter(({ bIndex }) => !usedB.has(bIndex));
  const removed = a.headers
    .map((aHeader, aIndex) => ({ aHeader, aIndex }))
    .filter(({ aIndex }) => !usedA.has(aIndex));

  const reordered = pairs
    .map((pair, index) => ({ pair, aOrder: index }))
    .filter(({ pair, aOrder }) => {
      const bOrder = pairs
        .slice()
        .sort((left, right) => left.bIndex - right.bIndex)
        .indexOf(pair);
      return bOrder !== aOrder;
    })
    .map(({ pair }) => ({
      aHeader: pair.aHeader,
      bHeader: pair.bHeader,
      aColumn: pair.aIndex + 1,
      bColumn: pair.bIndex + 1,
    }));

  const ambiguous = [...ambiguousNames].sort().map((name) => ({
    name,
    aIndexes: (aByName.get(name) ?? []).map((index) => index + 1),
    bIndexes: (bByName.get(name) ?? []).map((index) => index + 1),
  }));

  // 3. Guesses for the leftovers — presented to the user, never applied here.
  const suggested: ColumnDiff["suggested"] = [];
  const remainingA = removed.slice();
  const remainingB = added.slice();
  const looseName = (header: string): string => normalizeHeader(header).replace(/[^a-z0-9]/g, "");

  for (const bColumn of added) {
    const target = looseName(bColumn.bHeader);
    let match = remainingA.find((aColumn) => looseName(aColumn.aHeader) === target && target !== "");
    let reason = "header names are identical apart from case and punctuation";
    if (!match) {
      match = remainingA.find((aColumn) => {
        const source = looseName(aColumn.aHeader);
        return source !== "" && (source.includes(target) || target.includes(source));
      });
      reason = "one header name looks like part of the other";
    }
    if (!match) {
      const index = remainingB.indexOf(bColumn);
      match = remainingA[index];
      reason = "same position among the unmatched columns — a positional guess";
    }
    if (match) {
      suggested.push({
        bHeader: bColumn.bHeader,
        bIndex: bColumn.bIndex,
        aHeader: match.aHeader,
        aIndex: match.aIndex,
        reason,
      });
      remainingA.splice(remainingA.indexOf(match), 1);
      remainingB.splice(remainingB.indexOf(bColumn), 1);
    }
  }

  return { pairs, added, removed, reordered, ambiguous, suggested };
}

/** Suggest a key column. The suggestion is only a guess — the caller must confirm it. */
export function suggestKeyColumn(
  a: SheetTable,
  b: SheetTable,
  pairs: ColumnPair[],
): KeySuggestion | null {
  interface Candidate {
    header: string;
    nameBonus: boolean;
    unique: boolean;
    score: number;
  }
  const candidates: Candidate[] = [];
  for (const pair of pairs) {
    const valuesA = a.rows.map((row) => row[pair.aIndex] ?? "");
    const valuesB = b.rows.map((row) => row[pair.bIndex] ?? "");
    const usable = (values: string[]): boolean => {
      const seen = new Set<string>();
      for (const value of values) {
        const trimmed = value.trim();
        if (trimmed === "") return false;
        const key = trimmed.toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
      }
      return true;
    };
    const unique = valuesA.length > 0 && valuesB.length > 0 && usable(valuesA) && usable(valuesB);
    const nameBonus = IDENTIFIER_RE.test(pair.aHeader.trim());
    const score = (unique ? 2 : 0) + (nameBonus ? 1 : 0);
    if (score > 0) candidates.push({ header: pair.aHeader, nameBonus, unique, score });
  }
  candidates.sort((left, right) => right.score - left.score);
  const best = candidates[0];
  if (!best) return null;
  const reason = best.unique
    ? best.nameBonus
      ? "the name looks like an identifier and every value is unique in both files"
      : "every value is unique in both files"
    : "the name looks like an identifier (values are not unique, so duplicates will be reported)";
  return { header: best.header, reason };
}

function keyMap(
  table: SheetTable,
  columnIndex: number,
  toggles: CompareToggles,
): Map<string, number[]> {
  const map = new Map<string, number[]>();
  table.rows.forEach((_, rowIndex) => {
    const key = normalizeKey(cellAt(table, rowIndex, columnIndex), toggles);
    const existing = map.get(key);
    if (existing) existing.push(rowIndex);
    else map.set(key, [rowIndex]);
  });
  return map;
}

function emptyTable(): SheetTable {
  return { headers: [], rows: [] };
}

/** Compute the full diff. Pure: the tables and options are never modified. */
export function diffSheets(a: SheetTable, b: SheetTable, options: DiffOptions = {}): DiffResult {
  const toggles: CompareToggles = { ...DEFAULT_TOGGLES, ...options.toggles };
  const columns = diffColumns(a, b, options.columnMapping ?? {});
  const warnings: Finding[] = [];
  const notes: Finding[] = [];

  const aPairsByHeader = new Map(columns.pairs.map((pair) => [normalizeHeader(pair.aHeader), pair]));
  const suggestion = suggestKeyColumn(a, b, columns.pairs);

  // ---- Key resolution -------------------------------------------------------
  let mode: DiffResult["mode"] = "none";
  let keyColumnIndex: number | null = null;
  let keyColumnBIndex: number | null = null;
  let keyHeader: string | null = null;
  let problem: string | null = null;

  const requested = options.keyColumn ?? null;
  if (options.positional) {
    mode = "positional";
  } else if (requested !== null && requested !== "") {
    const pair = aPairsByHeader.get(normalizeHeader(requested));
    if (!pair) {
      const inA = a.headers.some((header) => normalizeHeader(header) === normalizeHeader(requested));
      problem = inA
        ? `The key column "${requested}" has no matching column in file B, so rows could not be matched by key.`
        : `The key column "${requested}" was not found in file A.`;
      warnings.push({
        id: "key-missing",
        severity: "warning",
        title: "The chosen key column could not be used",
        details: [
          problem,
          "Pick another key column, or explicitly compare rows by position (a clearly-labeled fallback that can be wrong when rows were inserted or deleted).",
        ],
      });
    } else {
      mode = "key";
      keyHeader = pair.aHeader;
      keyColumnIndex = pair.aIndex;
      keyColumnBIndex = pair.bIndex;
    }
  }

  if (mode === "none") {
    warnings.push({
      id: "no-key",
      severity: "warning",
      title: "No key column selected — rows were not compared",
      details: [
        suggestion
          ? `Suggested key: "${suggestion.header}" — ${suggestion.reason}. This is only a guess and was not applied.`
          : "No column with unique values in both files was found.",
        "Choose a key column to match rows (id, email, SKU, …), or explicitly switch to positional comparison.",
      ],
    });
  }

  if (mode === "positional") {
    warnings.push({
      id: "positional",
      severity: "warning",
      title: "Comparing rows by position (explicit fallback)",
      details: [
        "Row 1 of file A was compared with row 1 of file B, and so on.",
        "This is a guess: as soon as a row is inserted or deleted, every following row is reported as changed. A key column is almost always more accurate.",
      ],
    });
  }

  // ---- Header quality warnings ---------------------------------------------
  const headerFindings: Finding[] = [];
  const ambiguousByFile = (indexes: number[], file: string): string | null => {
    if (indexes.length === 0) return null;
    return `${file}: column${indexes.length === 1 ? "" : "s"} ${indexes.join(", ")}`;
  };
  const emptyHeaders = (table: SheetTable, file: string): string | null => {
    const indexes = table.headers
      .map((header, index) => ({ header, index }))
      .filter(({ header }) => header.trim() === "")
      .map(({ index }) => index + 1);
    return indexes.length === 0 ? null : `${file}: column${indexes.length === 1 ? "" : "s"} ${indexes.join(", ")}`;
  };
  for (const entry of columns.ambiguous) {
    const parts = [
      ambiguousByFile(entry.aIndexes, "File A"),
      ambiguousByFile(entry.bIndexes, "File B"),
    ].filter((part): part is string => part !== null);
    headerFindings.push({
      id: `duplicate-header-${entry.name}`,
      severity: "warning",
      title: `Duplicate header "${entry.name}"`,
      details: [
        `${parts.join(" and ")} share the header "${entry.name}" (compared case-insensitively).`,
        "Those columns were left out of the comparison because it is impossible to tell which one should match which.",
      ],
    });
  }
  for (const [table, file] of [
    [a, "File A"],
    [b, "File B"],
  ] as const) {
    const empty = emptyHeaders(table, file);
    if (empty) {
      headerFindings.push({
        id: `empty-header-${file}`,
        severity: "warning",
        title: `${file} has a column with no header name`,
        details: [
          `${empty} has an empty header — usually a trailing delimiter or an extra column.`,
          "The column is shown with an empty name and is compared by position within its header name.",
        ],
      });
    }
  }
  for (const finding of headerFindings) {
    warnings.push(finding);
  }

  // ---- Row comparison -------------------------------------------------------
  const added: RowChange[] = [];
  const removed: RowChange[] = [];
  const changedCells: CellChange[] = [];
  const ambiguousKeys: AmbiguousKey[] = [];
  const moved: DiffResult["moved"] = [];
  const emptyKey = { a: [] as number[], b: [] as number[] };
  let matchedRows = 0;

  const compareMatchedRow = (rowA: number, rowB: number): void => {
    for (const pair of columns.pairs) {
      const oldValue = cellAt(a, rowA, pair.aIndex);
      const newValue = cellAt(b, rowB, pair.bIndex);
      if (valuesDiffer(oldValue, newValue, toggles)) {
        changedCells.push({
          key: keyHeader === null ? `row ${dataRowRecord(rowA)}` : normalizeKey(cellAt(a, rowA, keyColumnIndex ?? 0), toggles),
          aRecord: dataRowRecord(rowA),
          bRecord: dataRowRecord(rowB),
          column: pair.aHeader,
          aColumn: pair.aIndex + 1,
          bColumn: pair.bIndex + 1,
          oldValue,
          newValue,
        });
      }
    }
  };

  if (mode === "key" && keyColumnIndex !== null && keyColumnBIndex !== null) {
    const mapA = keyMap(a, keyColumnIndex, toggles);
    const mapB = keyMap(b, keyColumnBIndex, toggles);
    const keys = new Set<string>([...mapA.keys(), ...mapB.keys()]);

    for (const key of [...keys].sort()) {
      const rowsA = mapA.get(key) ?? [];
      const rowsB = mapB.get(key) ?? [];
      if (key === "") {
        emptyKey.a.push(...rowsA.map(dataRowRecord));
        emptyKey.b.push(...rowsB.map(dataRowRecord));
        continue;
      }
      const duplicateA = rowsA.length > 1;
      const duplicateB = rowsB.length > 1;
      if (duplicateA || duplicateB) {
        ambiguousKeys.push({
          key,
          aRecords: rowsA.map(dataRowRecord),
          bRecords: rowsB.map(dataRowRecord),
          reason:
            duplicateA && duplicateB ? "duplicate-in-both" : duplicateA ? "duplicate-in-a" : "duplicate-in-b",
        });
        continue;
      }
      const rowA = rowsA[0];
      const rowB = rowsB[0];
      if (rowA !== undefined && rowB !== undefined) {
        matchedRows += 1;
        const before = changedCells.length;
        compareMatchedRow(rowA, rowB);
        if (changedCells.length === before && rowA !== rowB) {
          moved.push({ key, aRecord: dataRowRecord(rowA), bRecord: dataRowRecord(rowB) });
        }
      } else if (rowB !== undefined) {
        added.push({ key, record: dataRowRecord(rowB), values: cellsAt(b, rowB) });
      } else if (rowA !== undefined) {
        removed.push({ key, record: dataRowRecord(rowA), values: cellsAt(a, rowA) });
      }
    }
  } else if (mode === "positional") {
    const count = Math.max(a.rows.length, b.rows.length);
    for (let index = 0; index < count; index += 1) {
      const hasA = index < a.rows.length;
      const hasB = index < b.rows.length;
      if (hasA && hasB) {
        matchedRows += 1;
        compareMatchedRow(index, index);
      } else if (hasB) {
        added.push({ key: null, record: dataRowRecord(index), values: cellsAt(b, index) });
      } else if (hasA) {
        removed.push({ key: null, record: dataRowRecord(index), values: cellsAt(a, index) });
      }
    }
  }

  // ---- Notes ----------------------------------------------------------------
  if (emptyKey.a.length > 0 || emptyKey.b.length > 0) {
    const details: string[] = [];
    if (emptyKey.a.length > 0) {
      details.push(`File A: ${emptyKey.a.length} row${emptyKey.a.length === 1 ? "" : "s"} with an empty key — record${emptyKey.a.length === 1 ? "" : "s"} ${describeRecords(emptyKey.a)}.`);
    }
    if (emptyKey.b.length > 0) {
      details.push(`File B: ${emptyKey.b.length} row${emptyKey.b.length === 1 ? "" : "s"} with an empty key — record${emptyKey.b.length === 1 ? "" : "s"} ${describeRecords(emptyKey.b)}.`);
    }
    details.push("Rows without a key value cannot be matched, so they were left out of added/removed/changed.");
    warnings.push({
      id: "empty-keys",
      severity: "warning",
      title: `Rows with an empty key: ${emptyKey.a.length + emptyKey.b.length}`,
      details,
    });
  }

  if (ambiguousKeys.length > 0) {
    const shown = ambiguousKeys.slice(0, MAX_FINDING_DETAILS);
    warnings.push({
      id: "ambiguous-keys",
      severity: "warning",
      title: `Duplicate keys: ${ambiguousKeys.length} value${ambiguousKeys.length === 1 ? "" : "s"} could not be matched`,
      details: [
        ...shown.map((entry) => {
          const parts: string[] = [];
          if (entry.aRecords.length > 0) parts.push(`File A record${entry.aRecords.length === 1 ? "" : "s"} ${entry.aRecords.join(", ")}`);
          if (entry.bRecords.length > 0) parts.push(`File B record${entry.bRecords.length === 1 ? "" : "s"} ${entry.bRecords.join(", ")}`);
          return `Key "${entry.key}": ${parts.join(" — ")}.`;
        }),
        ambiguousKeys.length > shown.length ? `…and ${ambiguousKeys.length - shown.length} more duplicate keys.` : "",
        "These rows were excluded from cell comparison: with a duplicated key it is impossible to tell which row should match which.",
      ].filter((detail) => detail !== ""),
    });
  }

  if (moved.length > 0) {
    const shown = moved.slice(0, MAX_FINDING_DETAILS);
    notes.push({
      id: "moved-rows",
      severity: "info",
      title: `Rows whose position changed but whose content did not: ${moved.length}`,
      details: [
        ...shown.map((entry) => `Key "${entry.key}": file A record ${entry.aRecord} is file B record ${entry.bRecord}.`),
        moved.length > shown.length ? `…and ${moved.length - shown.length} more.` : "",
      ].filter((detail) => detail !== ""),
    });
  }

  if (removed.length >= 10 && removed.length > added.length * 5) {
    notes.push({
      id: "swapped-hint",
      severity: "info",
      title: "These files might be swapped",
      details: [
        `File A has ${removed.length} rows that are not in file B, while file B has only ${added.length} new rows.`,
        "If you meant to compare the newer file against the older one, swap them. Nothing was changed automatically.",
      ],
    });
  }

  if (columns.pairs.length === 0) {
    warnings.push({
      id: "no-matching-columns",
      severity: "warning",
      title: "No column names matched between the two files",
      details: [
        "Headers are matched trimmed and case-insensitively. If the names really are different, map the columns manually.",
        "Without a single matching column, no cells could be compared.",
      ],
    });
  }

  const summary: DiffSummary = {
    aRows: a.rows.length,
    bRows: b.rows.length,
    matchedRows,
    addedRows: added.length,
    removedRows: removed.length,
    changedRows: new Set(changedCells.map((cell) => `${cell.aRecord}:${cell.bRecord}`)).size,
    changedCells: changedCells.length,
    ambiguousKeys: ambiguousKeys.length,
    emptyKeyRowsA: emptyKey.a.length,
    emptyKeyRowsB: emptyKey.b.length,
  };

  return {
    mode,
    toggles,
    key: { mode, column: keyHeader, suggested: suggestion, problem },
    columns,
    summary,
    added,
    removed,
    changedCells,
    ambiguousKeys,
    emptyKey,
    moved,
    warnings,
    notes,
  };
}

/** Convenience wrapper for callers that have no tables at all. */
export function emptySheet(): SheetTable {
  return emptyTable();
}
