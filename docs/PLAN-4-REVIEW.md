# Plan 4 review (Codex GPT-6.1-Sol): review, fix test-first, prove

You are a Herdr delegate. Do the work yourself; no agent_bridge, no other agents. Never run git (the lead commits).
Read first: `docs/PLAN-4-DATA-ENGINE.md` (the design and the rules of section 0), then `docs/PLAN-4-RESTART.md`
(sandbox facts: one MSBuild node, the vitest shim, `-Suites`, what only the lead can run).

## Situation
Three Codex Terra lanes built plan 4 in parallel; the lead merged them into `main` (this folder). Lane A: store and
ingestion (`ExcelApp/Data/Store/**`, `ExcelApp/Data/Ingest/**`, `DataCommands.cs`, `ExportChild.cs`). Lane B: query,
automation, formulas (`ExcelApp/Data/Query/**`, `ExcelApp/Formula/**`, `QueryCommands.cs`). Lane C: the Data interface
(`ExcelApp/ui/src/data/**`). Their answers are in `.lane-answers/` (one file per lane): read the deviations and the
"not proven" lists first; they are your starting leads.

## You own now
Every file the three lanes own (plan section 3) and their tests. Not: `Contracts.cs`, `StoreRoot.cs`, `DataHost.cs`,
`contracts.ts`, `CachePolicy.cs`, `Session.cs`, `MainForm.cs`, `Program.cs`, `Bridge.cs`, `EditSet.cs`, `EditJournal.cs`,
`check.ps1`, `ExcelApp.Tests/Program.cs` (lead). A needed change there: write it in your answer, with the exact diff.

## Review, in this order
1. **Rule 0 compliance** (plan section 0), proven by grep and reading, each with the command in your answer:
   - no `File.Write*`, `File.Create`, `new FileStream(` ... for writing, `File.Move`, `File.Delete`, `Directory.Create*`
     outside `StoreRoot`/`ConfigRoot` in lane code (tests may create made-up folders under `D:\ExcelAppTest\syn*` only);
   - no read of a company file except through the `--build` child; no ZIP/XML/CSV reader on `.xlsx/.xls/.csv`;
   - no `Log.Write`/`Console` of cell values, sheet, column or file names (look at every call);
   - every Excel child closes only its own workbook and app, never kills Excel, has `DisplayAlerts=false`;
   - nothing in `%TEMP%`.
2. **Contract match**: every command of plan section 4 is registered, takes the listed args and returns the listed
   shape (compare with `ExcelApp/ui/src/data/contracts.ts` field by field, camelCase). Every event in the plan is emitted
   with the `contracts.ts` shape. List each mismatch and fix it on the lane side.
3. **Correctness**: the version model (2.4: latest, locked, as-of, rollback, post-close), the safety rules (2.7), the
   diff pairing of duplicate keys, typing (leading zeros never lost), period and lock arithmetic, joins (nulls never
   match), the query engine (nulls last, case-insensitive text), formula semantics against Excel's documented behaviour.
4. **Concurrency**: the watcher, the work queue, the hourly lock timer, the scheduler, `Loaded` handlers in lane B
   (views, links, alerts, dashboards) running on background threads; the `ResultStore`; the catalog and
   `versions.json` locks. Look for races, unobserved task exceptions, deadlocks, handlers that throw into the engine.
5. **Error paths**: a broken `.xvt`, a half-written `.part`, a missing column, an empty sheet, a recipe that fails
   validation, an export child that times out, a full disk. Every user-facing error is a plain sentence (`UserError`).
6. **Speed**: re-run the speed lines; all targets of plan section 6 must print and hold.
7. **Interface against the contract**: `mockData.ts` returns the same shapes as the host.

## How to fix
Test first: write the failing test, see it fail, fix, see it pass. Keep fixes small and in the owning file.
Checks you can run in the sandbox: `powershell -NoProfile -File check.ps1 -Offline -Suites store,ingest,query,automation,formula,journal`;
in `ExcelApp/ui` with `$env:NODE_OPTIONS='--require=./sandbox-shim.cjs'`: `pnpm exec tsc --noEmit`,
`pnpm exec vitest run --pool=threads --configLoader runner`, `pnpm exec vite build --configLoader runner`.
Live Excel tests, the app walk-through and screenshots are the lead's: list what you want run.

## Answer
`.lane/review-answer.md`: line 1 DONE or BLOCKED; then findings as a table (severity, file:line, what, fixed how, test);
the grep commands and their results for rule 0; the check counts and SPEED lines; what you could not prove and what the
lead must run live. Reply in the pane with the path only.
