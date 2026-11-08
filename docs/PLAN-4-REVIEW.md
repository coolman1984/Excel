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

## Leads from the lead's verification (2026-11-08), check these besides the order above
Baseline on main (lead, live): viewer 439 passed; ExcelApp 610 passed; tsc 0; vitest 40/40; build.ps1 OK.
1. **Formula engine (lane B) is under-proven.** The live 60-formula comparison (`FormulaTests.Live`) uses only 12
   arithmetic patterns. Widen it to every function of the plan's list (text, dates on the 1900 system incl. the
   1900-02-29 quirk, lookups incl. XLOOKUP match/search modes, IF/IFERROR, error values #DIV/0! #N/A #VALUE! #REF!,
   empty-cell semantics, text-number coercion). Absolute references (`$A$1`) are parsed but dropped from the AST: keep
   them. XLOOKUP binary search modes are rejected: implement or return a plain-sentence error. Write each live case so
   it runs when `offline == false` (the lead runs it) and the offline case list grows to the plan's 150+ real cases.
2. **Sandbox storage** keeps edited columns in store JSON; the plan wants `TableFile` (`.xvt`). Move it.
3. **Recipe run status** is not persisted by the host, so the Recipes page shows "Status unavailable" until an event.
   Persist the last run (time, ok, message) in the recipe's config and return it from the recipe list.
4. **xvt write 10M x 10** varied 1.4 to 9.2 s across runs against the 10 s target; find the variance (GC, buffer size,
   disk flush) and make it hold with margin.
5. **Lane A live tests** are written in a dense one-line style; make them readable without changing what they prove.
6. **Lane B's request** to add `history` to the wire `QuerySpec` is declined: questions use the history commands.
7. **Interface**: the Ingestion page scrolls the whole shell (top bar cut off); "Last load" shows a raw ISO time
   (format it like the other pages); the main bundle is 680 kB (lazy-load the Data mode).

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
