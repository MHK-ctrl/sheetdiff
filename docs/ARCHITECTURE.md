# Architecture

## Goals

- Answer one question: **what exactly changed between these two sheets?**
- Be private by construction: neither file ever leaves the browser tab.
- Keep the comparison engine deterministic and unit-testable.
- Never modify, normalize or "repair" the user's data.

## Non-goals

- Evaluating formulas (only stored cell values are read).
- Fuzzy-matching keys or auto-repairing rows.
- Encoding detection beyond UTF-8 + BOM.
- Any server-side component.

## Data flow

```
file A (drop / picker) ──► ArrayBuffer ──► UTF-8 decode (BOM preserved) ─┐
paste event ─────────────► text ─────────────────────────────────────────┤
.xlsx ───────────────────► ArrayBuffer ──► read-excel-file (stored values only)
                                                                         ▼
                                              parseCsvText / parseXlsxBuffer   ← pure
                                                                         │
                                                            SheetTable {headers, rows}
                                                                         │
                                                                         ▼
                        diffSheets(tableA, tableB, {keyColumn, toggles, columnMapping, positional})
                                                                         │
                     ┌───────────────────────────────────────────────────┤
                     ▼                                                   ▼
        render summary + tables (DOM)                  buildMarkdownReport(result, a, b, meta)
        (textContent only — no innerHTML)                       │
                                                      copy to clipboard / .md download
```

## Modules

| File | Responsibility |
| --- | --- |
| `src/diff.ts` | Pure engine. Column matching, key resolution, row matching, cell comparison, duplicate-key handling, warnings and notes. No DOM, no network. |
| `src/parse.ts` | File → `SheetTable` normalization. CSV/TSV/semicolon/pipe via Papa Parse, `.xlsx` via `read-excel-file`, size guard, BOM and line-ending detection, ragged-row reporting. |
| `src/report.ts` | Pure Markdown export. Takes a `DiffResult` plus both parsed sheets and metadata (file names, timestamp, repo URL). |
| `src/main.ts` | DOM wiring only: file/paste intake, key menu, toggles, manual mapping, rendering, clipboard/download. |

## Key decisions

- **Rows are matched by a key column, never by position — unless the user explicitly asks for it.** Positional comparison is available as a clearly-labeled fallback, and it warns that inserted or deleted rows shift everything below them. A silent positional fallback would produce thousands of false "changed" rows.
- **Duplicate keys are reported and excluded.** When a key value appears more than once, it is impossible to tell which row should match which, so those rows are dropped from cell comparison and listed as ambiguous with their record numbers. The tool never picks the first occurrence.
- **A suggested key is a labeled guess.** Suggestions are ranked by uniqueness first and by identifier-like names second (`id`, `email`, `sku`, …), shown in the UI with the reason, and only applied when the user selects them.
- **Column matching is exact (trimmed, case-insensitive).** Anything left over is offered as a *suggestion* with a stated reason — including the honest "same position among the unmatched columns — a positional guess". Suggestions are never applied automatically.
- **Manual mappings are applied before automatic matching**, so a deliberate choice is never overridden. Automatic matching skips duplicated header names, because it is impossible to tell which of two columns with the same name should match which; those columns are reported and excluded, and the user can map them by hand (the manual mapping resolves the first occurrence of that name).
- **Comparison toggles are off by default, and the report records which ones were on.** Toggle order is: trim → case-fold → numeric → empty/missing. A missing cell is *not* the same as an empty string unless `emptyEqualsMissing` is enabled. The report always shows both raw values, so a reader can see exactly what was compared.
- **The engine never rewrites values it is given.** Toggles only change the comparison; the values stored in `changedCells` are the raw ones from the files.
- **CSV record numbers are parsed record numbers.** A newline inside a quoted field means record number ≠ editor line number. This is stated in the UI and in the report, because a wrong row number is worse than no row number.
- **Delimiter detection is always called a guess.** Papa Parse's auto-detection is good but not infallible, so each file has its own manual override and auto results are labeled as guesses. Detection runs on a blank-skipping pass, because Papa Parse scores candidates by averaging field counts — a trailing newline (which parses as one extra empty record) would otherwise skew the score and break detection for ordinary two-column files.
- **XLSX reading** uses `read-excel-file` (MIT, no `eval`, browser-friendly, bundles its own XML parser). Only stored cell values are read; formulas are not evaluated, and a formula cell shows its last saved value. `trim: false` is passed so the reader does not trim cell strings — trimming is a toggle, not a file property. The reader drops trailing empty rows and columns, so xlsx record numbers are parsed-record numbers, exactly like CSV.
- **BOM handling is read-only.** Files are decoded with `TextDecoder("utf-8", { ignoreBOM: true })` so a leading U+FEFF stays visible to the parser, which reports it and then ignores it. Nothing is ever written back to disk.
- **Data is never trusted as HTML.** Sheet-derived strings (headers, cell values, error messages) are inserted with `textContent` only.
- **5 MB cap per file.** Both files are parsed in the tab; the cap keeps the UI responsive. The check happens before decoding.
- **`base: "./"` in `vite.config.ts`.** GitHub Pages serves the app from a repository subpath (`/sheetdiff/`), so all asset URLs must be relative.
- **No runtime network access.** There is no `fetch`, `XMLHttpRequest`, analytics snippet, or remote asset anywhere in `src/`.

## Testing

- `tests/diff.test.ts` runs in Node via vitest against the pure engine: identical files, added/removed rows with record numbers, changed cells with both raw values, missing cell vs empty string, duplicate keys, missing key column, empty keys, moved rows, swapped-file hint, positional fallback, column added/removed/reordered, manual mapping, duplicate and empty headers, every toggle, key normalization and key suggestions.
- `tests/parse.test.ts` covers CSV parsing (quoted commas and newlines, delimiters, manual override, BOM, blank lines, ragged rows with parsed record numbers, unclosed quotes, mixed line endings, empty input, size limit), a committed `.xlsx` fixture read through the same browser entry point the app uses, and report rendering.
- `tests/fixtures/people.xlsx` is generated by `tests/fixtures/make-xlsx-fixture.mjs` (a tiny ZIP writer with no extra dependencies) and committed, so the test suite does not depend on running the generator.
- `tests/setup.ts` supplies a `DOMParser` global so the xlsx reader runs on its browser code path in Node (browsers provide it natively). This keeps tests on the production path instead of a test-only module swap.
- Manual smoke test: `npm run dev`, load `examples/before.csv` and `examples/after.csv`, confirm every category renders and the report downloads.

## Deployment

- `.github/workflows/deploy.yml` runs on every push to `main`: `npm ci` → `npm test` → `npm run build` (which type-checks with `tsc --noEmit`) → `configure-pages` → `upload-pages-artifact` → `deploy-pages`.
- Pages uses the **GitHub Actions** build type (not branch-based), so the deployed site is exactly the tested artifact.
- The workflow needs `pages: write` and `id-token: write`; `configure-pages` is invoked with `enablement: true` so a fresh repo can enable Pages on its first run.

## Failure modes

- **Wrong key column** produces nonsense added/removed reports. The suggested key is labeled as a guess, and duplicate-key detection reports what cannot be matched.
- **Ragged rows** produce missing cells; they are reported with record numbers and compared as missing, never repaired.
- **Quoted newlines shift record numbers**: documented in the UI and the report.
- **Formula cells** show their last saved value, which may be stale if the file was written by a tool that does not cache results.
- **Non-UTF-8 input** is decoded with replacement characters; a note is shown if any U+FFFD appears. The tool does not attempt to guess the original encoding.