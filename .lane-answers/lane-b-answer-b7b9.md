# Lane B: B7–B9

Implemented recipes and scheduling, sandbox and library in `ExcelApp/Data/Query/ReportRecipes.cs`, `Scheduler.cs`, `Sandbox.cs`, `Library.cs`; updated `QueryCommands.cs` and `ExcelApp.Tests/AutomationTests.cs`. Alerts and scheduler are registered in `DataServices.Owned`.

Offline lane check: **50 passed, 0 failed**, build 0 warnings and 0 errors. October-to-November recipe replay, missed-run persistence, sandbox compare/keep, and a library export test that excludes synthetic table data pass. SPEED: group-by 10M 0.693s; filter/sort 1M 0.418s; normalized join 1M × 1M 2.539s; sandbox create/discard 1M 0.004s.

Integration limitations requiring lead action:
- Lane A's `TableFile` is absent in this worktree. The sandbox adapter persists edits as store JSON, not the planned edited-column `.xvt`; connect it to the merged table-file implementation. A sandbox based on a moving report reference currently follows that reference when re-opened.
- `ConfigRoot` cannot access a user-selected path outside its own folder. Library `.exlib` import/export is therefore restricted to that folder to obey the file-root rule. A lead-owned approved external configuration-file access contract is needed for the file-picker flow.
- Pack sheets use an optional `sheets` field on a pack step; the current UI contract does not yet describe that field.
