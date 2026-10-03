import "./style.css";
import {
  diffSheets,
  normalizeHeader,
  type CellChange,
  type DiffResult,
  type Finding,
  type RowChange,
} from "./diff";
import {
  detectFormat,
  fileSizeError,
  parseCsvText,
  parseXlsxBuffer,
  utf8ByteLength,
  type ParsedSheet,
  type SheetFormat,
} from "./parse";
import { buildMarkdownReport, formatBytes } from "./report";

const REPO_URL = "https://github.com/MHK-ctrl/sheetdiff";

/** Detail rows rendered in the page before a "…and N more" note. */
const MAX_UI_DETAIL_ROWS = 200;

function el<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing element #${id}`);
  return node as T;
}

type Side = "a" | "b";

interface SideElements {
  dropZone: HTMLDivElement;
  fileInput: HTMLInputElement;
  delimiterSelect: HTMLSelectElement;
  info: HTMLParagraphElement;
}

const sides: Record<Side, SideElements> = {
  a: {
    dropZone: el("drop-a"),
    fileInput: el("file-a"),
    delimiterSelect: el("delimiter-a"),
    info: el("info-a"),
  },
  b: {
    dropZone: el("drop-b"),
    fileInput: el("file-b"),
    delimiterSelect: el("delimiter-b"),
    info: el("info-b"),
  },
};

const statusEl = el<HTMLParagraphElement>("status");
const controlsEl = el<HTMLElement>("controls");
const keySelect = el<HTMLSelectElement>("key-select");
const keyHint = el<HTMLParagraphElement>("key-hint");
const positionalToggle = el<HTMLButtonElement>("positional-toggle");
const toggleTrim = el<HTMLInputElement>("toggle-trim");
const toggleCase = el<HTMLInputElement>("toggle-case");
const toggleEmpty = el<HTMLInputElement>("toggle-empty");
const toggleNumeric = el<HTMLInputElement>("toggle-numeric");
const mappingArea = el<HTMLElement>("mapping-area");
const mappingList = el<HTMLElement>("mapping-list");
const summaryEl = el<HTMLElement>("summary");
const resultsEl = el<HTMLElement>("results");
const reportActionsEl = el<HTMLElement>("report-actions");
const copyButton = el<HTMLButtonElement>("copy-report");
const downloadButton = el<HTMLButtonElement>("download-report");

interface LoadedFile {
  fileName: string;
  format: SheetFormat;
  /** CSV text (already decoded, BOM preserved for reporting). */
  text: string | null;
  /** Raw bytes for .xlsx. */
  buffer: ArrayBuffer | null;
}

let loaded: Record<Side, LoadedFile | null> = { a: null, b: null };
let parsed: Record<Side, ParsedSheet | null> = { a: null, b: null };
let result: DiffResult | null = null;
let keyColumn: string | null = null;
let positional = false;
let columnMapping: Record<string, string> = {};
let generatedAt = "";
let lastError: string | null = null;

function setStatus(message: string, kind: "" | "ok" | "error" = ""): void {
  statusEl.textContent = message;
  statusEl.className = kind === "" ? "status" : `status status-${kind}`;
}

function currentToggles() {
  return {
    trimWhitespace: toggleTrim.checked,
    ignoreCase: toggleCase.checked,
    emptyEqualsMissing: toggleEmpty.checked,
    numericAware: toggleNumeric.checked,
  };
}

function decodeUtf8(buffer: ArrayBuffer): string {
  // ignoreBOM keeps a leading U+FEFF in the string so the parser can report it.
  return new TextDecoder("utf-8", { ignoreBOM: true }).decode(buffer);
}

async function loadInto(side: Side, file: File): Promise<void> {
  const sizeError = fileSizeError(file.size);
  if (sizeError) {
    loaded[side] = null;
    parsed[side] = null;
    setStatus(sizeError, "error");
    renderSide(side);
    await refreshAfterLoad();
    return;
  }
  const format = detectFormat(file.name);
  const buffer = await file.arrayBuffer();
  loaded[side] = {
    fileName: file.name,
    format,
    text: format === "csv" ? decodeUtf8(buffer) : null,
    buffer,
  };
  columnMapping = {};
  await parseSide(side);
}

async function loadPastedText(text: string): Promise<void> {
  const sizeError = fileSizeError(utf8ByteLength(text));
  if (sizeError) {
    setStatus(sizeError, "error");
    return;
  }
  const side: Side = loaded.a === null ? "a" : "b";
  loaded[side] = { fileName: "pasted text", format: "csv", text, buffer: null };
  columnMapping = {};
  await parseSide(side);
}

async function parseSide(side: Side): Promise<void> {
  const file = loaded[side];
  if (!file) {
    parsed[side] = null;
    renderSide(side);
    await refreshAfterLoad();
    return;
  }
  try {
    if (file.format === "xlsx") {
      parsed[side] = await parseXlsxBuffer(file.buffer ?? new ArrayBuffer(0), { fileName: file.fileName });
    } else {
      const manual = sides[side].delimiterSelect.value;
      parsed[side] = parseCsvText(file.text ?? "", {
        fileName: file.fileName,
        delimiter: manual === "" ? undefined : manual,
      });
    }
    lastError = null;
  } catch (error) {
    parsed[side] = null;
    lastError = `Could not read "${file.fileName}": ${error instanceof Error ? error.message : String(error)}`;
  }
  renderSide(side);
  await refreshAfterLoad();
}

function renderSide(side: Side): void {
  const elements = sides[side];
  const file = loaded[side];
  const sheet = parsed[side];
  if (!file) {
    elements.info.textContent = "No file loaded.";
    elements.dropZone.classList.remove("has-file");
    return;
  }
  elements.dropZone.classList.add("has-file");
  const formatHint = file.format === "xlsx" ? "Excel workbook" : `${sheet?.stats.delimiterLabel ?? "?"} delimiter`;
  if (!sheet) {
    elements.info.textContent = `${file.fileName} — could not be read.`;
    return;
  }
  const rows = sheet.table.rows.length;
  const columns = sheet.table.headers.length;
  elements.info.textContent = `${file.fileName} — ${rows} row${rows === 1 ? "" : "s"}, ${columns} column${columns === 1 ? "" : "s"}, ${formatBytes(file.format === "xlsx" ? sheet.stats.byteSize : sheet.stats.byteSize)} (${formatHint})`;
}

function keyOptions(): string[] {
  const sheetA = parsed.a;
  if (!sheetA) return [];
  const headersB = new Set((parsed.b?.table.headers ?? []).map(normalizeHeader));
  return sheetA.table.headers.filter((header) => headersB.has(normalizeHeader(header)));
}

function rebuildKeySelect(): void {
  const options = keyOptions();
  const previous = keyColumn;
  keySelect.replaceChildren();
  const none = document.createElement("option");
  none.value = "";
  none.textContent = "Don't compare rows (choose a key column)";
  keySelect.append(none);
  for (const header of options) {
    const option = document.createElement("option");
    option.value = header;
    option.textContent = header === "" ? "(column with no header)" : header;
    keySelect.append(option);
  }
  keySelect.value = previous !== null && options.includes(previous) ? previous : "";
  if (keySelect.value === "") keyColumn = null;
}

function renderMapping(): void {
  if (!result) {
    mappingArea.hidden = true;
    return;
  }
  const suggestions = result.columns.suggested;
  if (suggestions.length === 0) {
    mappingArea.hidden = true;
    mappingList.replaceChildren();
    return;
  }
  mappingArea.hidden = false;
  mappingList.replaceChildren();
  for (const suggestion of suggestions) {
    const row = document.createElement("div");
    row.className = "mapping-row";
    const label = document.createElement("span");
    label.textContent = `“${suggestion.bHeader}” (file B) ↔`;
    const select = document.createElement("select");
    const none = document.createElement("option");
    none.value = "";
    none.textContent = "— not mapped —";
    select.append(none);
    for (const header of parsed.a?.table.headers ?? []) {
      const option = document.createElement("option");
      option.value = header;
      option.textContent = header === "" ? "(column with no header)" : header;
      select.append(option);
    }
    select.value = columnMapping[suggestion.bHeader] ?? "";
    select.addEventListener("change", () => {
      if (select.value === "") delete columnMapping[suggestion.bHeader];
      else columnMapping[suggestion.bHeader] = select.value;
      void recompare();
    });
    const apply = document.createElement("button");
    apply.type = "button";
    apply.className = "secondary";
    apply.textContent = "Apply guess";
    apply.addEventListener("click", () => {
      columnMapping[suggestion.bHeader] = suggestion.aHeader;
      select.value = suggestion.aHeader;
      void recompare();
    });
    const reason = document.createElement("span");
    reason.className = "reason";
    reason.textContent = `guessed: “${suggestion.aHeader}” — ${suggestion.reason}`;
    row.append(label, select, apply, reason);
    mappingList.append(row);
  }
}

const SEVERITY_NAMES = { error: "Error", warning: "Warning", info: "Note" } as const;

function findingArticle(finding: Finding): HTMLElement {
  const article = document.createElement("article");
  article.className = `finding severity-${finding.severity}`;
  const title = document.createElement("h3");
  const badge = document.createElement("span");
  badge.className = "badge";
  badge.textContent = SEVERITY_NAMES[finding.severity];
  title.append(badge, finding.title);
  article.append(title);
  const list = document.createElement("ul");
  for (const detail of finding.details) {
    const item = document.createElement("li");
    item.textContent = detail;
    list.append(item);
  }
  article.append(list);
  return article;
}

function chip(label: string, value: string, kind = ""): HTMLElement {
  const wrapper = document.createElement("div");
  wrapper.className = kind === "" ? "chip" : `chip chip-${kind}`;
  const labelEl = document.createElement("span");
  labelEl.className = "chip-label";
  labelEl.textContent = label;
  const valueEl = document.createElement("span");
  valueEl.className = "chip-value";
  valueEl.textContent = value;
  wrapper.append(labelEl, valueEl);
  return wrapper;
}

function renderSummary(diff: DiffResult): void {
  const { summary } = diff;
  summaryEl.replaceChildren(
    chip("Added rows", String(summary.addedRows), "added"),
    chip("Removed rows", String(summary.removedRows), "removed"),
    chip("Changed rows", String(summary.changedRows), "changed"),
    chip("Changed cells", String(summary.changedCells), "changed"),
    chip("Matched rows", String(summary.matchedRows)),
    chip("A rows", String(summary.aRows)),
    chip("B rows", String(summary.bRows)),
  );
  if (summary.ambiguousKeys > 0) {
    summaryEl.append(chip("Duplicate keys", String(summary.ambiguousKeys)));
  }
  summaryEl.hidden = false;
}

function tableWrap(table: HTMLTableElement): HTMLElement {
  const wrap = document.createElement("div");
  wrap.className = "table-wrap";
  wrap.append(table);
  return wrap;
}

function buildRowTable(
  changes: RowChange[],
  headers: string[],
  recordLabel: string,
  rowClass: string,
): DocumentFragment {
  const fragment = document.createDocumentFragment();
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  const recordHead = document.createElement("th");
  recordHead.className = "record";
  recordHead.textContent = recordLabel;
  headRow.append(recordHead);
  for (const header of headers) {
    const th = document.createElement("th");
    th.textContent = header === "" ? "(no header)" : header;
    headRow.append(th);
  }
  thead.append(headRow);
  table.append(thead);

  const tbody = document.createElement("tbody");
  for (const change of changes.slice(0, MAX_UI_DETAIL_ROWS)) {
    const tr = document.createElement("tr");
    tr.className = rowClass;
    const recordCell = document.createElement("td");
    recordCell.className = "record";
    recordCell.textContent = change.key === null ? `record ${change.record}` : `${change.record} · ${change.key}`;
    tr.append(recordCell);
    for (let index = 0; index < headers.length; index += 1) {
      const td = document.createElement("td");
      const value = change.values[index];
      td.textContent = value === null || value === undefined || value === "" ? "—" : value;
      tr.append(td);
    }
    tbody.append(tr);
  }
  table.append(tbody);
  fragment.append(tableWrap(table));
  if (changes.length > MAX_UI_DETAIL_ROWS) {
    const more = document.createElement("p");
    more.className = "more-rows";
    more.textContent = `…and ${changes.length - MAX_UI_DETAIL_ROWS} more — the exported report includes up to 200 per section.`;
    fragment.append(more);
  }
  return fragment;
}

function buildChangedTable(changes: CellChange[]): DocumentFragment {
  const fragment = document.createDocumentFragment();
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  for (const label of ["Key", "Column", "Old value", "New value", "A record", "B record"]) {
    const th = document.createElement("th");
    th.textContent = label;
    if (label.endsWith("record")) th.className = "record";
    headRow.append(th);
  }
  thead.append(headRow);
  table.append(thead);
  const tbody = document.createElement("tbody");
  for (const change of changes.slice(0, MAX_UI_DETAIL_ROWS)) {
    const tr = document.createElement("tr");
    const cells: { text: string; className?: string }[] = [
      { text: change.key },
      { text: change.column === "" ? "(no header)" : change.column },
      { text: change.oldValue ?? "—", className: "old-value" },
      { text: change.newValue ?? "—", className: "new-value" },
      { text: String(change.aRecord), className: "record" },
      { text: String(change.bRecord), className: "record" },
    ];
    for (const spec of cells) {
      const td = document.createElement("td");
      td.textContent = spec.text;
      if (spec.className) td.className = spec.className;
      tr.append(td);
    }
    tbody.append(tr);
  }
  table.append(tbody);
  fragment.append(tableWrap(table));
  if (changes.length > MAX_UI_DETAIL_ROWS) {
    const more = document.createElement("p");
    more.className = "more-rows";
    more.textContent = `…and ${changes.length - MAX_UI_DETAIL_ROWS} more — the exported report includes up to 200 per section.`;
    fragment.append(more);
  }
  return fragment;
}

function section(heading: string, body: Node | null, count: number): HTMLElement | null {
  if (body === null) return null;
  const wrapper = document.createElement("div");
  const h3 = document.createElement("h3");
  h3.textContent = `${heading} (${count})`;
  wrapper.append(h3, body);
  return wrapper;
}

function renderResults(diff: DiffResult): void {
  resultsEl.replaceChildren();
  resultsEl.classList.add("results");
  const heading = document.createElement("h2");
  heading.textContent = "Results";
  resultsEl.append(heading);

  const allClear =
    diff.summary.addedRows === 0 &&
    diff.summary.removedRows === 0 &&
    diff.summary.changedCells === 0 &&
    diff.warnings.every((warning) => warning.id === "no-key" || warning.id === "positional");
  if (allClear) {
    const message = document.createElement("p");
    message.className = "all-clear";
    message.textContent =
      diff.mode === "none"
        ? "No rows were compared yet — choose a key column above."
        : "The two files match: no added rows, no removed rows, no changed cells.";
    resultsEl.append(message);
  } else {
    const sections = [
      section("Added rows", diff.added.length > 0 ? buildRowTable(diff.added, parsed.b?.table.headers ?? [], "B record", "row-added") : null, diff.added.length),
      section("Removed rows", diff.removed.length > 0 ? buildRowTable(diff.removed, parsed.a?.table.headers ?? [], "A record", "row-removed") : null, diff.removed.length),
      section("Changed cells", diff.changedCells.length > 0 ? buildChangedTable(diff.changedCells) : null, diff.changedCells.length),
    ].filter((entry): entry is HTMLElement => entry !== null);
    if (sections.length === 0) {
      const message = document.createElement("p");
      message.className = "all-clear";
      message.textContent = "No added, removed or changed rows were found.";
      resultsEl.append(message);
    }
    for (const entry of sections) resultsEl.append(entry);
  }

  if (diff.columns.added.length > 0 || diff.columns.removed.length > 0 || diff.columns.reordered.length > 0) {
    const columnSection = document.createElement("div");
    const h3 = document.createElement("h3");
    h3.textContent = "Column changes";
    columnSection.append(h3);
    const list = document.createElement("ul");
    for (const column of diff.columns.added) {
      const item = document.createElement("li");
      item.textContent = `Added column: “${column.bHeader}” (file B, column ${column.bIndex + 1}).`;
      list.append(item);
    }
    for (const column of diff.columns.removed) {
      const item = document.createElement("li");
      item.textContent = `Removed column: “${column.aHeader}” (file A, column ${column.aIndex + 1}).`;
      list.append(item);
    }
    for (const entry of diff.columns.reordered) {
      const item = document.createElement("li");
      item.textContent = `Reordered column: “${entry.aHeader}” is column ${entry.aColumn} in A and ${entry.bColumn} in B.`;
      list.append(item);
    }
    columnSection.append(list);
    resultsEl.append(columnSection);
  }

  const problems = [...diff.warnings, ...parsed.a?.findings ?? [], ...parsed.b?.findings ?? []].filter(
    (finding) => finding.severity !== "info",
  );
  const notes = [...diff.notes, ...diff.warnings.filter((finding) => finding.severity === "info"), ...parsed.a?.findings ?? [], ...parsed.b?.findings ?? []].filter(
    (finding) => finding.severity === "info",
  );
  const seen = new Set<string>();
  const uniqueProblems = problems.filter((finding) => {
    const key = `${finding.id}:${finding.title}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  const seenNotes = new Set<string>();
  const uniqueNotes = notes.filter((finding) => {
    const key = `${finding.id}:${finding.title}`;
    if (seenNotes.has(key)) return false;
    seenNotes.add(key);
    return true;
  });

  if (uniqueProblems.length > 0) {
    const h3 = document.createElement("h3");
    h3.textContent = `Warnings (${uniqueProblems.length})`;
    resultsEl.append(h3);
    for (const finding of uniqueProblems) resultsEl.append(findingArticle(finding));
  }
  if (uniqueNotes.length > 0) {
    const h3 = document.createElement("h3");
    h3.textContent = `Notes (${uniqueNotes.length})`;
    resultsEl.append(h3);
    for (const finding of uniqueNotes) resultsEl.append(findingArticle(finding));
  }
  resultsEl.hidden = false;
}

/**
 * Recompute the diff from the current state and re-render. Called after any
 * file load, key change, toggle change or manual mapping.
 */
async function recompare(): Promise<void> {
  if (!parsed.a || !parsed.b) {
    result = null;
    renderAll();
    return;
  }
  result = diffSheets(parsed.a.table, parsed.b.table, {
    keyColumn,
    columnMapping,
    positional,
    toggles: currentToggles(),
  });
  generatedAt = new Date().toISOString();
  renderAll();
}

/** Load-triggered refresh: compares as soon as both files are parsed. */
async function refreshAfterLoad(): Promise<void> {
  if (parsed.a !== null && parsed.b !== null) {
    await recompare();
  } else {
    renderAll();
  }
}

function renderAll(): void {
  rebuildKeySelect();
  const ready = parsed.a !== null && parsed.b !== null;
  controlsEl.hidden = !ready;
  if (!ready) {
    summaryEl.hidden = true;
    resultsEl.hidden = true;
    reportActionsEl.hidden = true;
    result = null;
    if (lastError) setStatus(lastError, "error");
    else if (loaded.a === null || loaded.b === null) setStatus("Add both files to see the differences.");
    return;
  }
  if (!result) {
    // Both files parsed but the diff has not run yet (should be momentary).
    setStatus("Comparing…");
    return;
  }

  const { summary } = result;
  const excluded = summary.ambiguousKeys + summary.emptyKeyRowsA + summary.emptyKeyRowsB;
  if (result.mode === "none") {
    setStatus("Both files loaded. Choose a key column to match rows, or compare by position.", "");
  } else if (summary.addedRows === 0 && summary.removedRows === 0 && summary.changedCells === 0) {
    setStatus(
      excluded === 0
        ? "No differences found between the two files."
        : `No differences found in the rows that could be matched — but ${excluded} row${excluded === 1 ? " was" : "s were"} excluded (see warnings).`,
      excluded === 0 ? "ok" : "error",
    );
  } else {
    setStatus(
      `${summary.addedRows} added · ${summary.removedRows} removed · ${summary.changedRows} rows changed (${summary.changedCells} cells).`,
      "",
    );
  }
  if (result.key.suggested && keyColumn === null && !positional) {
    keyHint.textContent = `Suggested key: “${result.key.suggested.header}” — ${result.key.suggested.reason}. This is only a guess; pick it above to apply it.`;
  } else if (positional) {
    keyHint.textContent =
      "Comparing by position: row 1 of A against row 1 of B. This is a guess — inserted or deleted rows shift everything below them.";
  } else if (keyColumn !== null) {
    keyHint.textContent = `Rows are matched by “${keyColumn}”. Duplicate key values are reported and excluded, never silently picked.`;
  } else {
    keyHint.textContent = "";
  }
  renderSummary(result);
  renderResults(result);
  renderMapping();
  reportActionsEl.hidden = false;
}

function currentMarkdown(): string | null {
  if (!result || !parsed.a || !parsed.b || !loaded.a || !loaded.b) return null;
  return buildMarkdownReport(result, parsed.a, parsed.b, {
    aName: loaded.a.fileName,
    bName: loaded.b.fileName,
    generatedAt,
    repoUrl: REPO_URL,
  });
}

copyButton.addEventListener("click", () => {
  const markdown = currentMarkdown();
  if (markdown === null) return;
  navigator.clipboard.writeText(markdown).then(
    () => {
      copyButton.textContent = "Copied ✓";
      window.setTimeout(() => {
        copyButton.textContent = "Copy report";
      }, 1500);
    },
    () => setStatus("Clipboard access was blocked — use “Download .md” instead.", "error"),
  );
});

downloadButton.addEventListener("click", () => {
  const markdown = currentMarkdown();
  if (markdown === null) return;
  const blob = new Blob([markdown], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "sheetdiff-report.md";
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

for (const side of ["a", "b"] as const) {
  const elements = sides[side];
  elements.fileInput.addEventListener("change", () => {
    const file = elements.fileInput.files?.[0];
    if (file) void loadInto(side, file);
  });
  elements.dropZone.addEventListener("click", () => elements.fileInput.click());
  elements.dropZone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      elements.fileInput.click();
    }
  });
  elements.dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    elements.dropZone.classList.add("is-dragging");
  });
  elements.dropZone.addEventListener("dragleave", () => elements.dropZone.classList.remove("is-dragging"));
  elements.dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    elements.dropZone.classList.remove("is-dragging");
    const file = event.dataTransfer?.files?.[0];
    if (file) void loadInto(side, file);
  });
  elements.delimiterSelect.addEventListener("change", () => {
    if (loaded[side]) void parseSide(side);
  });
}

keySelect.addEventListener("change", () => {
  keyColumn = keySelect.value === "" ? null : keySelect.value;
  if (keyColumn !== null) positional = false;
  positionalToggle.textContent = positional ? "Comparing by position (click to use a key)" : "Compare by position instead";
  void recompare();
});

positionalToggle.addEventListener("click", () => {
  positional = !positional;
  if (positional) keyColumn = null;
  keySelect.value = keyColumn ?? "";
  positionalToggle.textContent = positional ? "Comparing by position (click to use a key)" : "Compare by position instead";
  void recompare();
});

for (const toggle of [toggleTrim, toggleCase, toggleEmpty, toggleNumeric]) {
  toggle.addEventListener("change", () => void recompare());
}

window.addEventListener("paste", (event) => {
  const target = event.target as HTMLElement | null;
  if (target && (target.tagName === "INPUT" || target.tagName === "SELECT" || target.tagName === "TEXTAREA")) {
    return;
  }
  const text = event.clipboardData?.getData("text/plain") ?? "";
  if (text === "") return;
  event.preventDefault();
  void loadPastedText(text);
});

renderAll();
