# Decision 3: ExcelApp measurements and outcomes

Date: 2026-11-07. Made-up files only (`D:\ExcelAppTest\syn-2m.xlsx`, `syn-20m.xlsx`, 20 million cells, 184 MB).

## Measurements (selftest, same machine, headless-style off-screen window)

| What | Target | Measured | Result |
|---|---|---|---|
| First screen from process start | 1,500 ms | 3,904 ms (page ready 3,632 ms; open to screen 246 ms) | missed |
| Jump to a far row, median / p95 / worst | 50 / - / 200 ms | 92 / 261 / 261 ms | missed |
| Filter one column (950,001 rows shown) | - | 630 ms | |
| Sort one column | - | 2,816 ms | |
| Find, nothing found, all sheets | - | 0.46 s | |
| Memory, host at start / after work | - | 71 / 221 MB | |
| Memory, WebView2 | - | 349 MB | |
| Save of 100 edits on syn-2m, save to reopened | - | 20.0 s (NASCA mark kept) | |
| First-ever open of syn-20m (Excel reads) | - | about 101 s | Excel-bound |

Reading the misses: 3.6 s of the first screen is the page becoming ready (dotnet start plus WebView2 start), not the data; the data arrives 246 ms after open. The jump numbers include the wait for the next painted frame, so they measure what the user sees. Not proven what dominates the page start; the next thing to try is a pre-warmed WebView2 environment and a smaller interface bundle.

## Style gate

Reading styles through Excel is off by default: the valid gate on a 20 x 200,000 workbook took 120 s (28,589 probes, partial, no errors), above the 90 s limit. Opt-in keeps a 200,000-probe / 120 s budget. A plain workbook adds 0.15 s and a 53-byte sidecar.

## Decisions

- Children are never killed; closing their stdin cancels them.
- `edge` takes and returns real sheet rows. The interface derives the real row as row number minus one, exact only when the sheet starts at row 1.
- Added bridge commands and parameters are additive to the contract.
- The policy scan allows the selftest to write PNG pictures; every other file write is checked.

## Open points for phase 2

Faster page start; jump latency; sort speed on 20 million cells; save speed (the child opens and rewrites the file in Excel); style gate speed; formulas calculation display.
