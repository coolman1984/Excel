# Lane B: B2 and B3 batch

Files: `ExcelApp/Data/Query/Engine.cs`, `Join.cs`, `Fuzzy.cs`, and `ExcelApp.Tests/QueryTests.cs`; progress recorded in `.lane/progress.md`. B1's `Pred.cs` and tests remain in the worktree.

Offline lane check: 32 passed, 0 failed; build 0 warnings, 0 errors.

- SPEED group-by 10M: 0.658s (target <1s).
- SPEED filter+sort 1M: 0.452s (target <1s; 500,000 selected rows, 1,000 distinct numeric values).
- SPEED messy join 1M x 1M: 2.749s (target <10s; normalized keys).

The group-by speed test exercises the optimized two-text-key sum/count path; other grouping shapes use a general implementation and have not been benchmarked at this scale. Similar matching uses the normalized composite key for multi-column joins. The 1M join benchmark uses normalized matching; a million-row workload with many unmatched similar candidates has not been benchmarked.
