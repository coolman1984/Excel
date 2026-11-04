# Excel Viewer

A small Windows program for looking at very large Excel files, smoothly, on an ordinary laptop.

- **The first time** a file is opened, a hidden, read-only Excel reads it (big files take a few minutes) and a fast copy is saved.
- **After that,** the file opens from the fast copy in under a second, with no Excel running.
- Scroll, search across all rows, and sort by any column. The source file is never changed.
- Copy cells with Ctrl+C or right-click Copy: the shown text, tab-separated, kept out of Windows clipboard history and cloud sync.
- Dates and money show as Excel shows them. A format that cannot be matched to Excel is marked with ⚠ instead of being guessed.

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
| `tools` | Earlier speed tests |
| `docs` | The design decisions |

History is in `HISTORY.md`.
