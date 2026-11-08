# Lane A handoff — incomplete

## Files changed or created
- `ExcelApp/Data/Store/TableFile.cs` (new)
- `ExcelApp/Data/Store/Catalog.cs` (new)
- `ExcelApp.Tests/StoreTests.cs` (replaced stub)
- `.lane/progress.md`, `.lane/answer.md` (new)

## Checks
- `python ExcelApp/run-guarded.py powershell -NoProfile -File check.ps1 -Offline -Suites store,ingest`: build succeeded; 5 passed, 0 failed. Four NU1900 warnings reported that the package vulnerability service SSL connection could not be established.
- No speed lines measured. The 10M × 10 write and header latency, 12-period union, pilot simulation, and live Excel tests are not proven.

## Status and deviations
- A1 is partial. Round trips cover all four column types, nulls, Unicode, empty and one-row tables, and an uncommitted part. Catalog version increment is covered. An explicit oversized input/header test and required performance tests are outstanding.
- A2–A13 were not implemented. The ingestion, workspace, exporter, commands and live tests remain stubs or absent. The lead must not treat this lane as ready to merge.
- The lazy columns implement the shared `Column` API, but are not concrete `NumberColumn`, `TextColumn`, or `BoolColumn` instances. Consumers that cast to concrete types will not work with these columns; integration and a suitable materialization or shared contract must be resolved and tested. The table writer also currently accepts only concrete columns, so cannot write a lazily opened table directly.
- Catalog locking is per instance, not shared across instances or processes; concurrent writers must be addressed before use.
- The input `Table.Rows` is an `int`, so a table with more than Int32.MaxValue rows cannot be represented by the current contract; malformed oversized headers still need a test.

## Contract requests
- Clarify a supported way to access materialized typed column arrays while retaining lazy loading; `Column` currently exposes only row-wise methods and `Table` is sealed.

## Problems / not proven
- No claim is made that the implementation meets the full XVT specification under malformed or concurrent input, or its speed target. Live tests require the lead's environment as stated in the restart instructions.
