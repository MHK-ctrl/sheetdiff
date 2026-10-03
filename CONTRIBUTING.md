# Contributing to SheetDiff

Thanks for wanting to help. This project has a few hard rules — they are what make it trustworthy:

1. **Comparison only.** Never add code that edits, repairs, or rewrites the user's data.
2. **No silent guesses.** A suggested key, a suggested column mapping, or a positional fallback must always be an explicit, labeled choice. If a decision can be wrong, the user makes it.
3. **No network at runtime.** No uploads, telemetry, analytics, remote fonts, or API calls. The files never leave the browser tab. If a feature needs a server, it does not belong in this project.
4. **UTF-8 only.** Do not claim encoding detection. UTF-8 (+BOM), stated plainly.
5. **`src/diff.ts` and `src/parse.ts` stay pure.** No DOM, no network, no globals — they must remain unit-testable in Node.
6. **Formulas are never evaluated.** Only values stored in the file are compared.
7. **Toggles stay off by default.** Any comparison that normalizes data must be opt-in and recorded in the report.

## Getting set up

```bash
git clone https://github.com/MHK-ctrl/sheetdiff.git
cd sheetdiff
npm install
npm run dev       # http://localhost:5173
npm test          # vitest
npm run build     # tsc --noEmit && vite build
```

## Making changes

- **Tests first for the engine.** Every new diff rule needs a unit test in `tests/diff.test.ts` (or `tests/parse.test.ts` for input handling) that pins its exact output — record numbers included.
- **Keep the diff small.** One concern per pull request.
- **Conventional commit messages** (`feat:`, `fix:`, `docs:`, …) are appreciated.
- **Bug reports:** include the Markdown report from the tool and, if possible, a minimal pair of files that reproduces the issue. Redact real customer data first — a hand-crafted example is just as good.

## Project layout

```
src/diff.ts      pure diff engine (column matching, row matching, cell comparison)
src/parse.ts     CSV/xlsx readers (pure with respect to the input)
src/report.ts    Markdown report builder (pure)
src/main.ts      DOM wiring, clipboard/download, size guard
tests/           vitest suites + fixtures
docs/            architecture notes
examples/        sample file pairs for trying the tool
```

## License

By contributing, you agree that your contributions are licensed under the [MIT license](LICENSE).