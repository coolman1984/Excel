# Lane B: B4–B6

Implemented formatting and page/release commands; saved views, exception counts, links, drill and lineage; and load-triggered and hourly alerts. Files changed: `ExcelApp/Data/Query/Format.cs`, `Views.cs`, `Drill.cs`, `Alerts.cs`, `Pred.cs`, `ExcelApp/Data/QueryCommands.cs`, `ExcelApp.Tests/QueryTests.cs`, and `ExcelApp.Tests/AutomationTests.cs`.

Offline lane check: **41 passed, 0 failed**, build 0 warnings and 0 errors. The synthetic load fires an alert within one evaluation; a repeated value is suppressed. SPEED: group-by 10M 0.765s; filter/sort 1M 0.534s; normalized join 1M × 1M 2.673s.

Link results are recomputed when run rather than cached; exception counts refresh on relevant load events. Multi-column null-group drill uses an internal `isNull` predicate. Alert prior values and last fired values are stored under `alerts/state.json` in the store.

Integration request: the lead-owned `DataHost.Dispose` should dispose lane B's alert service when lifecycle ownership is added; currently the timer and ingestion subscription live for the host process lifetime. Alert rule errors are isolated so other rules can evaluate; no per-rule error is currently exposed to the interface.
