# History

## 2026-11-07

- **The editor, planned:** the owner approved option A, a desktop app with a C# host and a WebView2 interface that looks like Excel, reads through Excel COM into the existing fast copy, filters and edits on the copy, and saves through Excel so files stay NASCA-protected. The complete one-shot build plan for an implementing agent is `docs\PLAN-3-EXCELAPP.md`. Nothing was built yet.

## 2026-11-04

- **First thinking:** measured Excel COM on protected files. Reading one screen takes about 2 ms at any size; opening is the slow part. Rust cannot link on this PC, and C# builds with nothing to install. See `docs\DECISION-1.md`.
- **Cache design:** one column format in memory and on disk, a reader that exists only while needed, and a disk switch off by default. Everything Excel saves here gets the NASCA mark, including CSV; files written by other programs do not. See `docs\DECISION-2-CACHE.md`.
- **Approval:** the owner, acting as IT manager, approved the disk copy for this laptop's local disk only.
- **Real-file test:** reopening a large protected file from the copy took a fraction of a second instead of minutes. The figures stay in private notes on this PC.
- **The viewer:** built `ExcelViewer` in C# Windows Forms: open, scroll, search, sort, and switch sheets.
- **Dates and money:** the copy stores each column's number format, and our display is checked against Excel's own text while the copy is made. Month and day names come from Excel, because Windows and Excel can differ (Sept and Sep). A format that does not match is marked ⚠ and never guessed.
- **Memory stop:** the owner set the stop at 96% Windows commit, one setting in `Machine.StopAt`. The older speed tools in `tools` still use their own lower limits.
- **Lessons from the real file:** sheets can hide zeros, Excel drops the minus when a negative rounds to zero, and Excel gives blank text for cells in hidden columns. All three are handled and tested on made-up data.
- **A hang, fixed:** an old copy with the same version mark failed to load and a hidden error box waited for 19 minutes. The version mark now changes with the layout, a damaged copy is read again, and unattended runs never show a box.
- **Copy:** the owner approved copying for their work on this laptop. Ctrl+C and right-click Copy give the shown text with number formats, tab-separated. The clipboard data asks Windows to leave it out of clipboard history and cloud sync. Both saved copies stay, and hidden rows and columns stay shown, by the owner's decision.
