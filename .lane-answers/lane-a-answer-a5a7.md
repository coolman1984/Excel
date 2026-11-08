# Lane A batch A5–A7

Offline check: 56 passed, 0 failed. XVT 10M × 10 write 1.354 s; header open 4.745 ms. Warm union 12 × 1M 0.469 s. Build reported four NU1900 vulnerability-service SSL warnings. The two-month pilot and listed synthetic ingestion cases pass.

Created or changed: `ExcelApp/Data/Store/{TableFile,Catalog,ExcelReader,WorkQueue,Typing,Versions,TableProvider}.cs`, `ExcelApp/Data/Ingest/{Periods,Compare,Recipes,Watcher,Recognizer,Ingestion}.cs`, `ExcelApp.Tests/{StoreTests,IngestTests}.cs`, `.lane/progress.md`, `.lane/answer.md`.

A5–A7 are implemented and the offline suite is green. The real Excel child and watcher timing have not been live-verified; the lead runs live checks per restart instructions. Cleanup reversal was tested for an explicit uppercase fix; all other combinations of fixes and concurrent watcher/reader interactions need further integration coverage. A8–A13 remain outside this batch. The updated A9/A12 contracts noted by the lead still need to be applied in those steps.
