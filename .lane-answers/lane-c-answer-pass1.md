# Lane C handoff — incomplete

The lane has not met every acceptance condition. The interface components and mock are in place, and the checks below pass, but the remaining work listed here must be completed before treating C1–C16 as done.

## Files

Existing C1/C2 work: `ExcelApp/ui/src/App.tsx`, `ExcelApp/ui/src/mockHost.ts`, `ExcelApp/ui/src/data/DataApp.tsx`, `ExcelApp/ui/src/data/mockData.ts`, `ExcelApp/ui/src/styles/data.css`, `ExcelApp/ui/src/data/ResultGrid.tsx`.

Added: `ExcelApp/ui/src/data/DataPage.tsx`, `SpecBuilder.tsx`, `ManagementPages.tsx`, `IngestionPage.tsx`, `OtherPages.tsx`, `data.test.ts`. Progress is in `.lane/progress.md`. No shared contract file was edited.

## Checks

- TypeScript `tsc --noEmit`: passed, exit 0.
- Vitest: 6 files, 31 tests passed, exit 0 (includes the existing 25 tests and 6 lane tests).
- Vite production build: passed, exit 0; warning for a chunk larger than 500 kB.
- Commands were run with the sandbox shim and `run-guarded.py`; no new packages or git commands were used.
- No speed lines apply to this UI lane. Editor switch preservation and real-host operation have not been exercised interactively.

## Incomplete work and deviations

- C3: paging, sorting, drill, lineage and actions exist, but row-status coloring depends on `_status` reaching the grid. The contract says reserved columns are excluded. The paging test checks the generation predicate rather than an asynchronous component race.
- C4: the builder supports a single basic filter, one grouping/aggregate, one computed expression and one join; advanced predicates accept JSON. It does not yet offer a complete visual Pred editor, multiple aggregates/computed fields/joins, or all requested editing and validation behavior.
- C5: Edit opens the view spec in Data; it does not save edits back to the same view.
- C6: Wizard covers seven steps and basic validation; suggested fixes lack before/after examples, header-row changes do not recompute the preview, the live stage timeline and queue are simplified, and recipe management is incomplete. It has not been tested through real Excel.
- C7: diff display cannot reliably construct result metadata from the current `DiffResult`. The current adapter guesses columns/count; trend uses period totals only. Further implementation and verification are required.
- C8–C9: Alert and recipe editors are simplified; recipe pack does not yet configure several sheets and validation/error handling needs completion.
- C10: pin requires an existing dashboard; dashboard creation and pin interaction need review.
- C11: direct grid editing is wired, but sandbox compare currently displays counts rather than a paged diff. Synthetic mock edits are kept in memory but compare is illustrative.
- C12: selection and literal preview are wired, but mock preview does not extract literals and import results are summarized only.
- C13: settings contract exposes only allowPdf/startWithWindows/maxParallelExcel; paths and separate PDF authorization are unavailable. UI displays unavailable paths when absent.
- C14–C15: mock covers the command names but its query engine and question parser are illustrative; it does not reproduce all 20 query patterns or every state transition. This is not a complete behavioral mock.
- C16: tests do not cover every required predicate round-trip and full wizard interaction; visual inspection and editor mode-switch behavior remain unproven. The restart explicitly assigns screenshots to the lead because Chrome cannot start in the sandbox.

## Contract requests

1. Return `ResultInfo` metadata (columns and row count) for `data.history.diff` and `data.sandbox.compare`, or provide a command to obtain it from `resultId`.
2. Expose status values for diff coloring while keeping lineage fields private; current ResultInfo contract excludes all `_` columns.
3. Expose read-only store/config paths and a separate PDF authorization capability in `data.settings.get` if Settings is to show them accurately.

## Screenshots for lead

At 1400×860, use `http://localhost:5173/?mode=data&page=` with each of: `home`, `data`, `views`, `ingestion`, `history`, `alerts`, `recipes`, `sandbox`, `library`, `settings`. Also capture Data with an open result and Ingestion with the wizard open, for twelve states. The latter two require interaction after navigating to the corresponding page. Inspect all pictures and report defects.

## Problems and evidence

The first attempted check used a path relative to the wrong working directory and `pnpm` without its Windows executable, producing path-not-found / FileNotFoundError; it did not alter source. Corrected checks used Node entry points through the guarded runner. The final checks above passed. Visual and real-host correctness are not proven; the restart says the lead performs screenshots and the integrated host walk-through.
