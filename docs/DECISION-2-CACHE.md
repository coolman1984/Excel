# Decision 2: the fast copy (cache)

Date: 2026-11-04. Numbers here come from **made-up data only**, measured on this PC with `tools\cache-bench`. Real-file results are kept out of this repository on purpose.

**Result:** the design was then tested on a real protected file. Reopening from the fast copy took a small fraction of a second instead of minutes, and the copy read back identical. The figures stay in private notes on this PC. The answer is yes, worth building for looking and searching, while the disk approval holds.

## The rule at the top

**The fast copy holds company data without NASCA protection.**

| Where the copy lives | Status |
|---|---|
| **In memory**, erased on close | The default. Always allowed |
| **On this laptop's local disk** | One clearly marked switch, **OFF by default**. Approved on 2026-11-04 by the owner, acting as IT manager, for **this laptop's local disk only** |
| Anywhere else | **Never.** No repository, chat, network share, cloud or sync folder. Nothing leaves this machine |

Two more things never change:
- The only reader of a protected file is Excel through COM. The tool never strips, avoids or defeats NASCA on the source.
- The switch (`--disk-cache-approved`) is refused for repository folders, OneDrive and any drive that is not a fixed local disk. These three refusals and the OFF default were each tested. `.gitignore` also blocks cache files from commits.

### Two honest limits of "in memory"
- Windows can move memory to its page file on disk, and hibernation writes memory to disk too ([Raymond Chen, Microsoft](https://devblogs.microsoft.com/oldnewthing/20071106-00/?p=24573)). "Erased on close" means the program lets go of it; it is not a promise that no trace ever touched the disk.
- Deleting a cache file removes it from view, as with any deleted file; it does not wipe the disk. Whether this disk is encrypted (BitLocker) could not be read from this account.

## The design in plain words

1. **The viewer** is small and always on. **The reader** is a separate small program that exists only while it is needed.
2. **Opening a file:** if a valid fast copy exists (and the switch is on), the viewer uses it and Excel never starts. Otherwise the viewer starts the reader.
3. **The reader** starts its own hidden, read-only Excel, then sends rows to the viewer through a private pipe, in blocks.
4. **First rows:** they appear about 50 to 90 ms after Excel finishes opening. Before that nothing can be shown, because only Excel can read the file. The viewer says so honestly.
5. **While the rest loads:** the screen stays usable. If the person jumps to a place not yet read, that block goes to the front of the queue (median 41 ms, worst 56 ms).
6. **When all rows are in:** the reader closes its Excel and ends. This is the add-on idea: at 6 million cells, Excel's 308 MB goes away and the viewer holds the data in 140 MB.
7. **If the viewer closes early:** the pipe breaks, the reader notices, and it closes its own Excel. Tested: closed with 109 of 300 blocks read, and Excel left in 1.5 s. Excel is never killed.
8. **Memory guard:** the reader checks Windows commit before every block. It waits at 85% and stops at 90%, and deletes any half-written cache. The owner can turn the wait off for one run (`--stop-only-at-90`); the stop at 90% stays on. The viewer itself, built later, has no wait and stops at 96%, as the owner decided. Excel itself can hold commit above 85% once a big file is open, so check memory before opening too. Once Excel is opening a file it cannot be interrupted, because it is never killed.

### Search and sort
- They run in the background on the viewer's copy, not in Excel, so Excel's "one request at a time" no longer blocks the screen. They can be cancelled, and they cover what has loaded so far, with progress shown.
- **Search** scans the packed text. **Sort** never moves data: it builds a list of row numbers, and the grid reads through that list.

### The grid
The Windows Forms `DataGridView` in virtual mode is Microsoft's official way: the grid asks for a cell only when it is about to be shown, and a just-in-time cache holds pages of rows ([virtual mode](https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/virtual-mode-in-the-windows-forms-datagridview-control), [just-in-time loading](https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/implementing-virtual-mode-jit-data-loading-in-the-datagrid)). Sort and edit are ours to handle. Keep rows shared, use full-row selection, and turn automatic sizing off ([scaling guide](https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/best-practices-for-scaling-the-windows-forms-datagridview-control)).

### Export for reports
Export goes back through Excel: the reader opens its hidden Excel, writes the chosen rows into a new workbook, and saves. Everything Excel saves here came out protected (see below), so a report keeps its NASCA protection. The viewer never writes its own Excel or CSV file, because that would come out plain.

### How numbers look (solved in the viewer: see `HISTORY.md`)
`Value2` gives raw values, so a date arrives as a number ([Microsoft](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)). The formatted text property works one cell at a time in practice ([Microsoft](https://learn.microsoft.com/en-us/office/vba/api/excel.range.text)), and reading cell by cell is far too slow. Plan: read each column's number format once and apply it in the viewer. Not built or measured. Colours and formulas are not shown either.

## Format chosen and why

**One format, used both in memory and on disk: our own column chunks.**

Every 1,000 rows (fewer on very wide sheets), each column is stored as plain numbers, or as text bytes packed together, with repeated text stored once. A small "kinds" list is added only where a column mixes numbers and text. Each chunk never changes once ready, so the screen reads it with no lock. On disk, a small directory at the end of the file says where everything is.

**In plain words:** it is the same idea as Arrow (data stored by column, ready to use with no unpacking), trimmed to what Excel needs.

Why not the real thing:
- **Excel columns are not one type.** One column can hold a number, then text, then an error. Arrow refused a mixed plain column in my test (`ArrowInvalid`). Arrow's answer is a "union" column, which costs 5 bytes more per value in its dense form ([Arrow spec](https://arrow.apache.org/docs/format/Columnar.html)).
- **Apache.Arrow for .NET needs a download** ([NuGet](https://www.nuget.org/packages/Apache.Arrow): 23.0.0, Apache-2.0, 578 KB), so it was skipped, as instructed. Real Arrow was tested with the `pyarrow` 25.0.1 already on this PC, in Python, so its timings are not a like-for-like comparison.
- **Parquet with Zstd** is much smaller on disk, but must unpack before showing anything: 34 ms against about 1 ms for a first screen. It suits archives better than a viewer. It is the choice to revisit if disk space becomes the main problem.

Real Arrow is slightly smaller and sorted faster here. Our copy is not the best on every line; it is the one that fits mixed Excel cells with no download.

## Benchmark (made-up data)

6 million cells: 300,000 rows by 20 columns, numbers and short text. One run each, nothing written to disk for the memory columns.

| | **Our copy** | Keep COM's raw arrays | Real Arrow (Python) | Parquet + Zstd (Python) |
|---|---|---|---|---|
| Memory: the data itself | **60 MB** (10.5 bytes per cell) | 241 MB (42 bytes per cell) | 49 MB (8.6 bytes per cell) | 31 MB packed; 151 MB once unpacked |
| Memory: whole program | 140 MB | 261 MB | 379 MB (268 MB is Python itself) | 711 MB |
| First screen from Excel opening done | 63 ms | 85 ms | not run through COM | not run through COM |
| Later screens | **0.13 ms** | 0.16 ms | 1.0 ms | 34 ms for the first |
| Reading all rows, Excel start to finish | 23.0 s (17.1 s is Excel) | 26.7 s (18.5 s is Excel) | not run through COM | not run through COM |
| Search, nothing found | **7 ms** | 47 ms | 54 ms | not run |
| Sort by number | 113 ms | 158 ms | 55 ms | not run |
| Sort by repeated text | 372 ms | 754 ms | 270 ms | not run |
| Sort by unique text | 155 ms | 297 ms | 77 ms | not run |
| Size as a file | 59 MB | not a file format | 50 MB | 31 MB |

The same copy on disk, reopened in a new program with no Excel:

| | Made-up, 6 million cells |
|---|---|
| Build (Excel already open) | 5.0 s |
| **First screen, from program start** | **14.7 ms** |
| Random screen | 0.29 ms typical; 14 ms worst |
| Search, exact bytes / ignoring upper and lower case | 31 ms / 416 ms |
| Sort by number / by text | 103 ms / 189 ms |
| Reads back identical to what was built | Yes (checked by a fingerprint) |
| Bypassing the Windows file cache | 17 ms for the first bytes; 241 MB per second for all of it |

## Which formats come out protected

Tested on a small made-up workbook, so no real data ever sat in a plain file. The first 8 bytes of each file:

| Made by | File | First bytes | Result |
|---|---|---|---|
| Excel, new workbook | xlsx | `<## NASC` | **Protected** |
| Excel, from a protected workbook | xlsx, xlsb, CSV, CSV UTF-8, text | `<## NASC` | **Protected, every one** |
| Our program | own format | `XVCACHE1` | Plain |
| Our program | text | normal text | Plain |
| Our program | zip | `PK` | Plain |
| Python, `pyarrow` | Arrow IPC | `ARROW1` | Plain |
| Python, `pyarrow` | Parquet | `PAR1` | Plain |

**The pattern:** what Excel writes gets the NASCA mark, and what another program writes does not. All test copies were deleted after their first bytes were read.

## What we learned from other tools

These are their own claims, not checked by me. We take ideas, not code.

| Tool | What it does well | What we take |
|---|---|---|
| [falke87/large-csv-viewer](https://github.com/falke87/large-csv-viewer) (C++/Qt; source is "all rights reserved") | Maps the file, indexes rows in the background, shows first rows early, runs cancellable background jobs. Claims 20 million rows open in 4.1 s | First rows before the index ends; a list of row numbers for sorted views; cancellable jobs |
| [bgreenwell/xleak](https://github.com/bgreenwell/xleak) (Rust, Calamine, MIT) | Loads small files fully and large ones lazily, split at 1,000 rows | A simple size rule. Its reader, Calamine, cannot open NASCA files |
| [Danmoreng/simple-csv-viewer](https://github.com/Danmoreng/simple-csv-viewer) (C++, Win32, MIT) | A Windows list that asks for rows on demand; no numbers claimed | Confirms the virtual-list approach |
| [AG Grid viewport](https://www.ag-grid.com/javascript-data-grid/viewport/) | The grid asks for a row range, gets a small buffer around it, and shows blanks until data arrives; [draws only what is visible](https://www.ag-grid.com/javascript-data-grid/dom-virtualisation/) | Small buffer around the screen; never freeze while waiting |
| "Data Viewer" | I could not identify one tool by this name. The nearest I found are [ParqEdit](https://github.com/ooliJP/ParqEdit) and [Tableverse](https://github.com/tableverse-labs/tableverse), not studied | Nothing yet |

Formats: Feather V2 is exactly the Arrow IPC file format ([Arrow](https://arrow.apache.org/docs/python/feather.html)). Parquet is a column format built for storage, with compression ([Parquet](https://parquet.apache.org/docs/overview/)).

Real and checked on 2026-11-04: Apache.Arrow 23.0.0 (NuGet); pyarrow 26.0.1 is the latest (PyPI), and 25.0.1 is what runs here. No new package was installed.

## The disk-cache switch and its rule

- **Default:** OFF. Memory only.
- **On:** only with the owner's approval, and only for a fixed local disk on this laptop, in a folder only the owner's account and Windows can read.
- **Built from:** only what Excel reads through COM.
- **Fresh or stale:** the copy records the source file's size and time. If either changes, the copy is no longer used (the check is designed; not yet tested on a changed file).
- **Half-finished copies:** written under a temporary name and renamed only when complete; deleted if stopped.
- **Must never be committed or sent anywhere.** The file name, sheet names and values stay on this machine.
