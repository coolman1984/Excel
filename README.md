# Excel Viewer

A small Windows program for looking at very large Excel files, smoothly, on an ordinary laptop.

- **The first time** a file is opened, a hidden, read-only Excel reads it (big files take a few minutes) and a fast copy is saved.
- **After that,** the file opens from the fast copy in about a second, with no Excel running.
- Scroll, search across all rows, and sort by any column. The source file is never changed.
- Copy cells with Ctrl+C or right-click Copy: the shown text, tab-separated, kept out of Windows clipboard history and cloud sync.
- Dates and money show as Excel shows them. A format that cannot be matched to Excel is marked with ⚠ instead of being guessed.

## ExcelApp (the editor)

ExcelApp is a desktop editor that looks and behaves like Excel. A C# host opens a WebView2 window, the interface is a React grid in `ExcelApp\ui`, and the host talks to it with small JSON messages. Cells come from the same fast copy the viewer uses; edits live on top of the copy and are written back by a separate hidden Excel process, so saved files stay NASCA-protected.

**Run it:** `ExcelApp\Start_ExcelApp.bat` (builds on first use with `ExcelApp\build.ps1`; needs Node with pnpm for the interface build and the .NET 8 SDK). Windows will show the app in `ExcelApp\app\`.

**What it does:** open, scroll, select, type with Excel's rules (numbers, dates, percent, formulas as text), undo and redo, copy, cut, paste, fill, bold/italic/colour/number format, filter and sort one or many columns, find and replace, column widths, freeze panes, Save and Save As, close prompts, recent files.

**What it does not do:** it does not calculate formulas (a formula you type is saved to Excel, which calculates it); no charts, pictures, pivot tables or macros; no editing of the first open's slow part: the first open of a file still waits for Excel to read it, later opens use the copy. .NET 8 runs out of support on 2026-11-10; moving to .NET 10 is a one-line change in the project files.

**Tests:** `powershell -File check.ps1` (add `-Offline` to skip the tests that start Excel).
## The rule

The fast copy holds data **without** NASCA protection. The owner, acting as IT manager, approved it on 2026-11-04 for **one laptop's local disk only**. Nothing may leave that machine.

- Copies are written only to `D:\ExcelViewerCache`, locked to the owner's account. Every write goes through `ExcelViewer\Core\CachePolicy.cs`, which refuses any other place and any other PC.
- Excel through COM is the only reader of a protected file. NASCA on the source is never stripped or avoided.
- Never commit, send or sync a cache file, a picture of real data, or any figure from a real file.

## How to run

Double-click `ExcelViewer\Start_ExcelViewer.bat`, or `ExcelViewer\app\ExcelViewer.exe` once built. It needs Excel and .NET 8 or newer.

## How to check

`powershell -File check.ps1` builds everything and runs the tests, including a hidden Excel on a made-up workbook. `-Offline` skips Excel.

## Layout

| Folder | What it holds |
|---|---|
| `ExcelViewer\Core` | The fast copy, the reader, the number formats, search and sort |
| `ExcelViewer` | The window and the start script |
| `ExcelViewer.Tests` | The tests (no packages needed) |
| `ExcelApp` | The editor: host, bridge, edit set, filter engine, save-back, interface (`ui`) |
| `ExcelApp.Tests` | The editor's tests (host, edit, filter, save, style) |
| `tools` | Earlier speed tests |
| `docs` | The design decisions, `PLAN-3-EXCELAPP.md` (the build plan) and `DECISION-3-EXCELAPP.md` (what was measured) |

History is in `HISTORY.md`.
