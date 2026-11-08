# Lane B: B10–B13

Implemented dashboard values, cache and load refresh; the 20 question forms; lane B command registration; and the formula library. The library export now uses `ExternalLibrary`, pack sheets match the updated contract, and report sandboxes store an as-of reference.

Files: `ExcelApp/Data/Query/Dashboards.cs`, `Questions.cs`, `QuerySpec.cs`, `Expr.cs`, `Library.cs`, `Sandbox.cs`; `ExcelApp/Data/QueryCommands.cs`; `ExcelApp/Formula/Parser.cs`, `Evaluator.cs`, `DependencyGraph.cs`, `FormulaReader.cs`; `ExcelApp.Tests/AutomationTests.cs`, `FormulaTests.cs`.

Offline lane check: **256 passed, 0 failed**, build 0 warnings, 0 errors. SPEED: group-by 10M 0.889s; filter/sort 1M 0.604s; join 1M × 1M 3.530s; sandbox 1M 0.006s; formula ordering and evaluation 100k 0.449s. The 60-formula live Excel comparison and FormulaReader check are written under `!offline`; the lead must run them.

Integration request: add optional `history: { reportTypeId, period }` to the shared QuerySpec wire contract. The local history-question adapter uses the existing history commands to compare the latest two non-rolled-back versions. Sandbox edited columns still use store JSON pending the planned TableFile integration.

Formula compatibility is not fully established: the offline 150-case block mainly exercises arithmetic, with a smaller set of targeted function tests. Edge cases including Excel's 1900 date system and lookup match/search modes need broader Excel comparison; XLOOKUP binary search modes are currently rejected. Absolute-reference markers are parsed but not retained in the AST. The formula engine and question parser therefore need review and expanded validation before acceptance; the passing offline count does not prove full B13 compliance.
