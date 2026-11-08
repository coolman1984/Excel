# Lane A handoff

A1–A13 are implemented. Final guarded `check.ps1 -Offline -Suites store,ingest`: **87 passed, 0 failed**. Speed: XVT 10M × 10 write **9.211 s**, header open **4.065 ms**, warm 12 × 1M union **0.598 s**. One earlier header timing exceeded its target (60.960 ms); subsequent runs passed. Four NU1900 warnings report the vulnerability-service SSL connection failure.

Quality flags, history diff and settings contracts, workspace add/refresh, PDF gate, command registration, the ingestion pilot and undo for recipe and automatic cleanup kinds have offline checks. Quality results include all flagged rows; cards keep up to 20 samples. PDF remains disabled pending approval.

A13's made-up Excel workbook, corrected reload, export and readback tests are written under `!offline` but **were not run here**. The lead must run them live. Actual Excel COM export, template handling, multi-sheet reads, watcher timing, and the ten-minute timeout path remain unverified in this environment.

Changed files: `ExcelApp/Data/DataCommands.cs`, `ExportChild.cs`, `Store/{TableFile,Catalog,ExcelReader,WorkQueue,Typing,Versions,TableProvider,Quality,Workspace,Exporter}.cs`, `Ingest/{Periods,Compare,Recipes,Watcher,Recognizer,Ingestion,History}.cs`, `ExcelApp.Tests/{StoreTests,IngestTests}.cs`, `.lane/progress.md` and `.lane/answer.md`. No lead-owned contract files were edited. The lane-code scan found no direct `File.Write*` calls.
