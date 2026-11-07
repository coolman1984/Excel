# Plan 3: ExcelApp, the Excel-like editor for very large files

Date: 2026-11-07. Written by the planner (Claude Fable 5.1) for an implementing agent (Opus 5.5 medium or Sonnet 5.5 high). Approved by the owner on 2026-11-07: option A, a desktop app with a C# host and a WebView2 user interface, reading through Excel COM, saving through Excel COM.

**This plan is executed in one run, start to finish, with no questions back to the owner.** Every decision an implementer would normally ask about is already made below. Where a measurement decides something, the plan says what to measure and what each outcome means. If something cannot be done as written, do the nearest thing that keeps every rule in section 0, write what you did and why in `HISTORY.md`, and continue.

How to give this plan to the agent (paste as the whole prompt):

```
Read D:\WORK\Software Development\GitHub\Excel\docs\PLAN-3-EXCELAPP.md and execute it completely, step by step, in order, without stopping for questions. Run every check it names. Commit after each step with the message it gives. Finish with the report described in its last section. English only. Never print, log, screenshot or commit the contents of a real company file; use made-up data for every test and picture.
```

---

## 0. Rules that never bend

These come from the owner, the company and the two earlier decisions (`docs\DECISION-1.md`, `docs\DECISION-2-CACHE.md`). Read both files first.

1. **Excel through COM is the only reader and the only writer of workbook files.** The app never parses, writes or converts an xlsx, xlsm, xlsb, xls or CSV itself. No SheetJS, no ClosedXML, no OpenXML, no CSV export. A file Excel saves on this PC comes out NASCA-protected; a file we would write would not. The clipboard is the only other way data leaves the app.
2. **The fast copy (cache) lives only in `D:\ExcelViewerCache`** and is written only through `ExcelViewer\Core\CachePolicy.cs`. Never write cache data anywhere else. The owner's kept copy `real-test.xvc` is never touched.
3. **Never print, log, screenshot or commit real data.** Tests and pictures use made-up workbooks only. Log files hold timings, counts, error type names and made-up test paths, never a cell value, a sheet name or a real file name. The existing `Mask` and `Shape` helpers in `CacheBuilder.cs` show the standard.
4. **Excel is never killed.** It is closed with `Workbooks.Close(false)` and `Quit()`, in a `finally`, in a child process with a time limit. If it hangs, the user is told and the child is left to finish; the main window never freezes.
5. **Nothing waits on a box nobody can see.** Every unattended mode (`--build`, `--save`, `--selftest`, tests) writes a line and exits with a code; it never shows a `MessageBox`.
6. **Memory stop at 96% Windows commit** (`Machine.StopAt`), checked before opening Excel and during every read, exactly as `CacheBuilder` does now.
7. **No new install on this PC except packages.** NuGet from nuget.org and npm through the configured proxy are allowed and work. No Visual Studio, no SDK, no Rust linker, nothing on C: beyond package caches. Keep every build output on D:.
8. **English only** in code, comments, UI text, commits and reports. Plain words; no jargon in UI text.
9. **Do not touch `Browser Excel\`** (the owner's Univer and glide-data-grid experiments) except to add it to `.gitignore` as instructed in step 12. Do not change `ExcelViewer\MainForm.cs` or `ExcelViewer\Program.cs`; the old viewer keeps working unchanged. Changes to `ExcelViewer\Core` are additive only and `check.ps1` must still pass after each one.
10. **One Claude worker at a time** on this PC. Do not start other agents.

---

## 1. What already exists, and what you reuse

| File | What it gives you | How ExcelApp uses it |
|---|---|---|
| `ExcelViewer\Core\CachePolicy.cs` | `Folder`, `PathFor(source)` (hashed name), `Validate`, `OpenForWrite`, `ReplaceWith`, `Delete`, `ThisPcApproved` | All cache and sidecar writes. Extend the extension check to allow `.xvs` (step 1) |
| `ExcelViewer\Core\CacheFile.cs` | `CacheData` (Magic `XVCACHE4`, `Formats`, `Sheets`, `Load`, `HeaderMatches`, `IsFreshFor`), `SheetInfo` (`Rows`, `Cols`, `FirstRow`, `FirstCol`, `BlockRows`, `ColWidths`, `Hidden`, `HideZeros`, `FormatAt`), `SheetReader` (`Cell(row,col)` gives `CellView(Text, Kind, Color?, Warn)`, `Column`, `LoadColumn`, `RawText`, `Show`), `ColumnChunk`, `CellKind` | Read every cell shown. Do not change the file format or the Magic |
| `ExcelViewer\Core\CacheBuilder.cs` | `CacheBuilder.Build(source, cachePath, report, cancelled)` runs its own hidden Excel and writes the copy; `Machine.Commit()`, `Machine.StopAt`, `Letters(n)`, `ReadBlock`, the null-if-mixed `ReadRuns` trick for per-column properties | Refactor so the same code can run on an already open workbook (step 1), used after Save |
| `ExcelViewer\Core\Engines.cs` | `SearchEngine.Run` (parallel, hits in order, cap 10,000), `SortEngine.Sort` (stable, numbers then text then bools then errors, empties last), `CopyText.Build` | Find, sort, copy. Add a sort overload over a given row list (step 1) |
| `ExcelViewer\Core\NumFmt.cs` | Excel number format rendering, verified against Excel when the cache is built | Format edited numbers and user-chosen number formats |
| `ExcelViewer\MainForm.cs` lines 375 to 384 | `ClipboardSink`: puts text on the clipboard with the three flags that keep it out of clipboard history and cloud sync | Copy the same eight lines into ExcelApp |
| `ExcelViewer\Program.cs` `BuildMode` | The child-process pattern: stdout lines `OPENING`, `OPENED`, `FORMATS`, `READ nn`, `CHECKING`, `RESULT ...`; stdin closing means the parent is gone | Same pattern for `--build` and `--save` in ExcelApp |
| `ExcelViewer.Tests\Program.cs` | `Check(name, ok, detail)` style, `MakeWorkbook(path)` that creates a made-up workbook with Excel COM (formats, zeros, hidden column, a 300,000-row sheet), `Policy()` that scans source files for forbidden calls, `BadCopies()` | Copy the pattern; the new test project has its own smaller generator |
| `tools\com-bench` | `dotnet bench.dll make <path> <rows> <cols>` creates a made-up workbook of any size (20 million cells took 43 s to open in Excel) | The big performance test file |
| `check.ps1` | Builds and runs the old tests; `-Offline` skips Excel | Extend to also run the new tests |
| `.gitignore` | `*.xvc`, `bin/`, `obj/`, `app/`, `screenshots/`, `*.lnk` | Extend (step 12) |

Measured facts to design against (from the decision documents): one screen from the cache 0.1 to 0.3 ms; full search over 6 million cells 7 to 31 ms exact, 416 ms ignoring case; sort 100 to 400 ms; Excel opens about 2 seconds per million cells; `Range.Value2` on a block of 1,000 rows about 20 ms; reading one cell at a time 300 times slower than a block and never allowed.

Toolchain on this PC, verified 2026-11-07: .NET SDK 8.0.427 (build with it; `RollForward=LatestMajor` runs on the .NET 10 runtime that is installed), Node 25.2.1, pnpm 11.7.0 (npm registry through the configured proxy, already working), git 2.55, Excel 16.0.14334 at `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`, WebView2 runtime 157.0.4341.42, NuGet package `Microsoft.Web.WebView2` 1.0.2903.40 already in `%USERPROFILE%\.nuget\packages` (no download needed). C: has 4 GB free, D: 15 GB.

---

## 2. The design, fixed

### 2.1 Folder layout (new, beside the old viewer)

```
ExcelApp\
  ExcelApp.csproj              net8.0-windows, WinExe, WindowsForms, references ..\ExcelViewer\ExcelViewer.csproj,
                               PackageReference Microsoft.Web.WebView2 1.0.2903.40
  Program.cs                   modes: (none) window; --build; --save; --selftest; --dev <url>
  MainForm.cs                  one Form holding one WebView2, native file dialogs, close prompt
  Bridge.cs                    JSON message router: WebMessageReceived -> handler on a worker -> PostWebMessageAsJson
  Session.cs                   the open workbook: cache, per-sheet SheetReader, row map, filters, sort, edits
  Cells.cs                     CellView -> wire format; alignment; edited-cell rendering through NumFmt
  FilterEngine.cs              distinct values, conditions, Apply -> int[] rowMap (filter + multi-key sort)
  EditSet.cs                   value and style edits, undo/redo stacks, dirty flag
  SaveBack.cs                  child-process worker: hidden Excel, write edits, Save or SaveAs, rebuild cache in the same Excel
  StyleStore.cs                sidecar .xvs: per-column runs of fill, font colour, bold, alignment; hidden rows/cols; frozen panes
  Log.cs                       timings and counts only, to %LOCALAPPDATA%\ExcelApp\log.txt, 1 MB rolling
  Start_ExcelApp.bat           double-click start (like ExcelViewer\Start_ExcelViewer.bat)
  build.ps1                    pnpm install + vite build + copy to app\ui + dotnet publish to app\
  ui\
    package.json, pnpm-lock.yaml, vite.config.ts, tsconfig.json, index.html
    src\main.tsx, App.tsx, bridge.ts, rowStore.ts, theme.ts, keys.ts
    src\grid\Grid.tsx          glide-data-grid DataEditor, Excel theme, row-number column, selection, editing hooks
    src\ribbon\*.tsx           Ribbon, tabs Home / Data / View, File backstage
    src\bars\FormulaBar.tsx, SheetTabs.tsx, StatusBar.tsx
    src\dialogs\FilterMenu.tsx, SortDialog.tsx, FindReplace.tsx, GoTo.tsx, SaveProgress.tsx, ClosePrompt.tsx
    src\styles\*.css
  app\                         build output (gitignored): ExcelApp.exe, ui\ (vite dist)
ExcelApp.Tests\
  ExcelApp.Tests.csproj        console, references ExcelApp; no packages
  Program.cs                   Check() style tests; "offline" argument skips Excel
```

### 2.2 Process model

- **ExcelApp.exe (window)**: WinForms host, one `WebView2` filling the form. All data stays here in native memory: the cache reader, the row map, the edits. Nothing listens on a port.
- **ExcelApp.exe --build "<source>"**: child; runs `CacheBuilder.Build` (plus the style sidecar, step 10); stdout lines as the old viewer; exits when done; closes its Excel if stdin closes.
- **ExcelApp.exe --save**: child; reads one JSON document from stdin (source path, target path, edits), opens a hidden Excel, writes the edits, saves, then rebuilds the cache **from the same open workbook** so the user never waits for a second Excel open; stdout lines `OPENING`, `OPENED`, `WRITING nn`, `SAVING`, `SAVED`, then the build lines, then `RESULT ...`.
- **ExcelApp.exe --selftest "<made-up file>" "<png folder>"**: the window, off screen, scripted through `ExecuteScriptAsync`; prints timings and takes pictures; refuses any file whose name does not start with `syn`.
- **ExcelApp.exe --dev http://localhost:5173**: the window loads the Vite dev server instead of the built UI, for hot reload while working.

Child processes are started with `UseShellExecute=false`, `CreateNoWindow=true`, redirected stdin/stdout, and killed only if they are *our* process and the user cancels **before** `OPENED` (Excel not yet opening a file); after that they are asked to stop by closing stdin, exactly as `MainForm.CancelBuild` does today. Read that method before writing yours.

### 2.3 Message protocol between UI and host

Every request from the UI is one JSON object `{ "id": n, "cmd": "...", ... }` sent with `window.chrome.webview.postMessage`. Every reply is `{ "id": n, "ok": true, "result": ... }` or `{ "id": n, "ok": false, "error": "plain words for the user" }`. Host-initiated events are `{ "event": "...", ... }` with no id. The handler runs on a thread-pool thread; the reply is posted back on the UI thread with `BeginInvoke`. Requests on the same session are serialised through one `SemaphoreSlim(1)` except `getRows`, which only reads and may run in parallel (SheetReader is not thread safe: give each sheet one reader behind a lock, or one reader per request thread through `ThreadLocal`, as `SearchEngine` does).

| cmd | args | result |
|---|---|---|
| `ready` | | `{ version, recent: [paths], approved: bool }` |
| `pickFile` | | shows the native Open dialog; `{ path }` or `{ path: null }` |
| `open` | `path` | starts opening; returns `{ started: true }`. Progress arrives as events `opening` `{ phase, percent, seconds }` and `opened` `{ workbook }` or `openFailed` `{ error }` |
| `workbook` | | `{ name, path, sheets: [{ index, name, hidden, rows, cols, firstRow, firstCol, colWidthsPx: [], frozenRows, frozenCols }], dirty }` |
| `sheet` | `sheet` | makes it current; `{ viewRows, filterActive, sort: [{col, desc}], filters: {col: condition} }` |
| `rows` | `sheet, from, count` (view indexes) | `{ from, rows: [[rowNumber, [cell, cell, ...]]] }`, cell = `[text, kind, align, color?, warn?, style?]` (kind 0 empty 1 number 2 text 3 bool 4 error; align 0 left 1 right 2 centre); up to 500 rows per call |
| `cellDetail` | `sheet, row, col` (view index) | `{ address: "C12", raw: "what the formula bar shows", edited: bool }` |
| `setCell` | `sheet, row, col, input` (what the user typed) | `{ changed: [[row, col]], dirty: true }` |
| `clearCells` | `sheet, ranges: [[r1,c1,r2,c2]]` | same |
| `setStyle` | `sheet, ranges, style: { bold?, italic?, underline?, fontColor?, fill?, align?, numberFormat?, border? }` | same |
| `undo` / `redo` | `sheet` | `{ changed, dirty, canUndo, canRedo }` |
| `copy` | `sheet, ranges` | puts text on the clipboard with `ClipboardSink`; `{ cells: n }`; cap `CopyText.MaxCells` |
| `cut` | same | copy, then remember a pending cut; the next paste clears the source |
| `paste` | `sheet, row, col` | reads clipboard text, splits on `\r\n` / `\n` and `\t`, writes edits; `{ changed, dirty }`; cap 200,000 cells |
| `fill` | `sheet, from: [r1,c1,r2,c2], to: [r1,c1,r2,c2]` | fill handle: repeats the source block over the target; `{ changed }` |
| `filterValues` | `sheet, col, search?` | `{ values: [{ text, count }], blanks: n, truncated: bool }` top 10,000 distinct shown texts among rows that pass every *other* column's filter, sorted as Excel sorts (numbers ascending, then text A to Z, blanks last) |
| `setFilter` | `sheet, col, condition or null` | condition is `{ kind: "values", include: [texts], blanks: bool }` or `{ kind: "number", op, a, b? }` (op: eq ne gt ge lt le between, top10, aboveAverage, belowAverage) or `{ kind: "text", op, value }` (op: eq ne contains notContains begins ends); result `{ viewRows, filterActive }` |
| `clearFilters` | `sheet` | same |
| `setSort` | `sheet, keys: [{ col, desc }]` (empty clears) | `{ viewRows }` |
| `setHeaderRow` | `sheet, hasHeader` | whether row 1 is a header kept on top; `{ viewRows }` |
| `find` | `sheet or -1 for all, query, matchCase?` | starts `SearchEngine.Run`; events `findProgress` `{ percent }`, `findHits` `{ hits: [[sheet,row,col,text]] }` in batches of 200, `findDone` `{ total, capped }` |
| `replace` | `sheet, query, replacement, all: bool, at?: [row,col]` | makes edits in text cells; `{ changed, count }` |
| `stats` | `sheet, ranges` | `{ count, numbers, sum, average, min, max }` over the shown rows, computed in the host |
| `setColWidth` | `sheet, col, px` | remembered as a style edit (saved as `ColumnWidth` on save) |
| `setFreeze` | `sheet, rows, cols` | remembered; saved as `FreezePanes` on save |
| `save` | `saveAs: bool` | `saveAs` shows the native Save dialog first; then starts the child; events `saving` `{ phase, percent }`, `saved` `{ path, seconds }` or `saveFailed` `{ error }`; after `saved` the host reopens from the rebuilt cache and sends `opened` again |
| `close` | | if dirty, the UI already showed the prompt; closes the session |
| `recent` | | `{ recent }` |
| `openExternal` | `what: "folder"` | opens the file's folder in Explorer |
| `log` | `text` | UI timings into the host log (no cell text allowed; the host rejects strings longer than 200 chars) |

Events: `opening`, `opened`, `openFailed`, `saving`, `saved`, `saveFailed`, `findProgress`, `findHits`, `findDone`, `fileChanged` (the source changed on disk while open), `memory` `{ commitPercent }` every 5 s when above 85%.

### 2.4 The data model in the host

- **Row map.** For the current sheet, `int[] viewToRow` after filter and sort, or null when neither is active (identity). Header rows (`firstDataRow`, default 1 when `hasHeader`, Excel's default) always come first, in order, unfiltered.
- **Edits overlay.** `EditSet` is a `Dictionary<(int sheet, int row, int col), CellEdit>` where `CellEdit` holds `Kind`, `Num`, `Text`, `Formula` (the typed text when it starts with `=`), and the style delta. Reading a cell checks the overlay first; an edited number is rendered through the cell's number format by building a one-row `ColumnChunk` with `ColumnChunk.Build(1, ...)` and calling `SheetReader.Show(chunk, 0, row, col)`, so an edited price still shows as money. A formula edit shows the formula text in grey italics (the value is unknown until Excel saves) and the formula bar shows the text.
- **Undo/redo.** Two stacks of `Change` records `{ sheet, cells: [(row, col, before, after)] }`. One user action is one record (one paste is one undo). Cleared on save.
- **Typed input parsing**, in `CurrentCulture`: empty → clear; starts with `=` → formula; `TRUE`/`FALSE` (any case) → bool; `double.TryParse` with `NumberStyles.Any` → number; a percent such as `12.5%` → number 0.125 and, if the cell's format is General, style edit `numberFormat: "0.00%"`; a date that `DateTime.TryParse` accepts → Excel serial (days since 1899-12-30, or 1904-01-01 when `cache.Date1904`) and, if General, `numberFormat: "yyyy-mm-dd"` (Excel's local short date is what Excel would pick; ours is honest and sortable); everything else → text.
- **Filters** are evaluated on the shown text for `values` and `text` kinds, and on the raw number for `number` kind (text cells fail number conditions, as in Excel). Blanks are `Kind == Empty` or empty text.
- **Dirty** is true when `EditSet` is not empty or a style, width or freeze edit exists.

### 2.5 The user interface, fixed

Look: Excel for Microsoft 365, current (2025) design. Fonts: `Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif` at 11 pt in cells (14.67 px at 100% zoom), 9 pt in the ribbon. Colours: Excel green `#107C41` (title bar, active tab underline, selection border, Save button), hover `#0F703B`, selected cell fill `rgba(16,124,65,0.08)`, selected row/column header fill `#CAEAD8` with text `#0F703B`, header fill `#F3F3F3`, header text `#444`, grid lines `#E1E1E1`, cell text `#000`, frozen split line `#B0B0B0`, filtered row numbers blue `#1E5EFF` (Excel colours them blue), edited and unsaved cells get nothing special (as Excel), formula cells pending save show italics `#555`. Icons: `@fluentui/react-icons` (MIT), regular 20 px in the ribbon.

Layout, top to bottom:
1. **Title row** (34 px): app icon, file name, a disabled `AutoSave Off` pill, and on the right a window-less area (the OS title bar stays; do not hide it).
2. **Ribbon tabs**: `File`, `Home`, `Data`, `View`. No tab is shown for features that do not exist. Below the tabs the ribbon body (92 px) with grouped buttons and group labels like Excel; a collapse chevron at the right end hides the body (remembered).
   - **Home**: Clipboard (Paste, Cut, Copy), Font (name box read-only showing the sheet font, size box, Bold, Italic, Underline, Borders menu with None / All / Outside / Bottom, Fill colour with the Excel palette, Font colour), Alignment (Left, Centre, Right, Top, Middle, Bottom, Wrap off/on), Number (format dropdown: General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text; buttons `%`, `,`, increase and decrease decimals), Cells (Insert, Delete, Format: disabled in version 1 with tooltip "coming later"), Editing (AutoSum writes `=SUM(range above or left)` as a formula edit, Clear menu: Clear All / Clear Contents / Clear Formats, Sort & Filter menu, Find & Select menu: Find, Replace, Go To).
   - **Data**: Sort & Filter (Sort A to Z, Sort Z to A, Sort dialog with up to 3 keys, Filter toggle, Clear, Reapply), Data Tools (Remove Duplicates and Text to Columns disabled with the same tooltip), Outline nothing.
   - **View**: Show (Gridlines, Headings, Formula Bar check boxes), Zoom (Zoom dialog 50 to 200, 100%, Zoom to Selection disabled), Window (Freeze Panes menu: Freeze Panes at the active cell's column, Freeze First Column, Unfreeze; row freezing per section 2.6), Workbook Views nothing.
   - **File** opens the backstage (full window, green left rail): Home (recent list), Open (Browse... and recent), Save, Save As (Browse...), Info (path, size, last read time, cache size, "Protected by NASCA: yes" read from the first 8 bytes `<## NASC` of the source, never from contents), Close, and `Back`.
3. **Formula bar row** (28 px): Name box (shows `C12` or `C12:F40`; typing an address or `A1:B5` and Enter jumps and selects), `fx`, the formula box (shows raw text: the formula text for edited formulas, the raw number or text otherwise; editing and Enter commits like `setCell`; Esc cancels).
4. **Grid**: see 2.6.
5. **Sheet tabs row** (28 px): tabs for visible sheets, active tab white with green bottom border, a disabled `+`, right-click menu `Unhide...` listing hidden sheets to show for viewing. Left and right scroll arrows when tabs overflow.
6. **Status bar** (24 px): `Ready` / `Reading file... 37%` / `Saving...`; when a filter is active `n of m records found` (Excel's wording); selection stats on the right `Average: x  Count: n  Sum: s` (numbers only, from `stats`); zoom percentage and a slider 50 to 200.

Keyboard, Excel parity: arrows, Ctrl+arrows (jump to edge of data in the shown rows: host answers `edge` through `cellDetail`-like call or the UI scans loaded rows then asks the host for the far edge: implement as host command `edge { sheet, row, col, direction }`), Page Up/Down, Home, Ctrl+Home, Ctrl+End, Shift+arrows extend, Ctrl+Shift+arrows, Ctrl+A (selects all in the UI only; copy of all is capped), Enter moves down, Tab moves right, Shift+Enter up, F2 edits, Delete clears, Esc cancels, Ctrl+C/X/V, Ctrl+Z/Y, Ctrl+F Find, Ctrl+H Replace, Ctrl+G Go To, Ctrl+S Save, Ctrl+O Open, Ctrl+B/I/U, Ctrl+Shift+L Filter toggle, Alt+Down on a header cell opens its filter menu, Ctrl+Page Up/Down switches sheets, Ctrl+Mouse wheel zooms.

### 2.6 The grid, fixed

Use **`@glideapps/glide-data-grid`** (MIT, canvas, already proven on millions of rows; a copy of its repo is in `Browser Excel\glide-data-grid` for reading, but install the package from npm). Decisions:

- `rowMarkers="none"`. Column 0 is our own **row-number column** (frozen, 52 px, header fill, text right-aligned, blue when the row survives an active filter), showing the *original* Excel row number (`FirstRow + row`), as Excel does under a filter. Data columns start at index 1 and carry Excel letters (`CacheBuilder.Letters(FirstCol + k)`) as titles, `hasMenu: true` on the header row when Filter is on (the glide header menu icon becomes the AutoFilter arrow; draw a funnel badge with `drawHeader` when that column has a condition).
- `freezeColumns = 1 + frozenCols`. Frozen rows: glide freezes only the header. Implement "Freeze Top Row" and "Freeze Panes" rows by **drawing the frozen rows inside the header**: set `headerHeight = 20 * (1 + frozenRows)` and in `drawHeader` paint the letter row and then each frozen row's cells for that column (text from the row store, same fonts and alignment). Cap frozenRows at 10. If this proves unstable after one honest attempt, ship freeze columns only, say so in `HISTORY.md`, and move on.
- Column width: Excel units to pixels `Math.Round(w * 7 + 5)` at 100% zoom; row height 20 px. Zoom scales both and the font.
- Data: glide calls `getCellContent([col,row])` synchronously, so **`rowStore.ts`** keeps blocks of 200 view rows in a `Map<blockIndex, rows>`, an LRU of 400 blocks, and fetches with `rows` on `onVisibleRegionChanged` for the visible range plus two screens above and below, coalescing requests (one in flight per block). A cell not yet loaded returns `{ kind: GridCellKind.Loading }`; when a block arrives call `ref.updateCells` for its range. After any `changed` list from the host, invalidate those cells and `updateCells`. After `setFilter`, `setSort`, `undo`, `paste` and `opened`, clear the store and `ref.updateCells` for the visible range (or remount the grid with a new key when `viewRows` changed).
- Selection: `rangeSelect="multi-rect"`, `columnSelect="multi"`, `rowSelect="multi"`; clicking the row-number column selects the row; the name box mirrors the selection; the status bar asks `stats` (debounced 150 ms, only when the selection has at most 2 million cells, otherwise shows `Count` only from the host's cheap count).
- Editing: glide's overlay editor (`onCellEdited` → `setCell` with the typed text; use `TextCell` with `allowOverlay: true` so every cell edits as text); `onDelete` → `clearCells`; `fillHandle` on → `onFillPattern` → `fill`; `onPaste` returns false and we handle Ctrl+V ourselves through the host (`paste`) because glide's paste cannot see the clipboard flags; `getCellsForSelection` returns loaded cells only so glide's own copy is disabled (`keybindings.copy = false`) and Ctrl+C goes to the host `copy`.
- Theme object: `accentColor #107C41`, `accentLight rgba(16,124,65,0.12)`, `textDark #000`, `textHeader #444`, `bgHeader #F3F3F3`, `bgHeaderHasFocus #CAEAD8`, `bgHeaderHovered #E6E6E6`, `borderColor #E1E1E1`, `horizontalBorderColor #E1E1E1`, `headerFontStyle "600 12px"`, `baseFontStyle "14.67px"`, `fontFamily` as above, `cellHorizontalPadding 4`, `lineHeight 1.2`.
- Alignment per cell from the wire `align` field; `[Red]` style colours from `color`; `warn` cells keep their `⚠` text; bool and error centred.
- Styled cells (edited styles, and read styles once step 10 exists): custom `drawCell` paints fill, bold/italic/underline, font colour and borders before the text.

---

## 3. The steps

Do them in this order. Each step ends with its checks and one commit. Do not merge steps. Do not skip a check; if a check fails, fix it before moving on. Timings you measure go into `docs\DECISION-3-EXCELAPP.md` (made-up data only).

### Step 0: Preflight (10 minutes)

1. Read `README.md`, `HISTORY.md`, `docs\DECISION-1.md`, `docs\DECISION-2-CACHE.md`, and all files listed in section 1.
2. Run `powershell -File check.ps1` and record that it passes (it starts a hidden Excel on a made-up workbook). If it fails before you changed anything, record the failing check names in `HISTORY.md` and continue; your work must not add failures.
3. Confirm the toolchain lines of section 1 with `dotnet --version`, `node -v`, `pnpm -v`, and `Test-Path "$env:USERPROFILE\.nuget\packages\microsoft.web.webview2\1.0.2903.40"`.
4. Make the test workbooks once, in `D:\ExcelAppTest\` (create it; it is outside the repo): `cd tools\com-bench; dotnet build -c Release; dotnet bin\Release\net8.0-windows\bench.dll make D:\ExcelAppTest\syn-2m.xlsx 100000 20` and `... syn-20m.xlsx 1000000 20`. Check the exact output folder and the `make` argument order in `tools\com-bench\Program.cs` first. These files are made up; keep them out of the repo.

No commit.

### Step 1: Additive changes to `ExcelViewer\Core` (1 hour)

1. **`CacheBuilder`**: split `Build` into `Build(source, cachePath, report, cancelled)` (unchanged signature and behaviour: starts Excel, opens, calls the next method, closes) and a new public `BuildFromOpenWorkbook(object excelApp, object workbook, string source, string cachePath, Action<string> report, Func<bool> cancelled, Action<object, object, CacheData> afterRead = null)` holding everything from "Size of each sheet" to `ReplaceWith`. The sampler thread and the memory stop stay. `afterRead` runs after the values are written and before the directory, while Excel is still open; step 10 uses it. The `FileInfo` for `SourceLength`/`SourceTicks` must be read **inside** `BuildFromOpenWorkbook` (after a save the file is new).
2. **`SortEngine`**: add `Sort(CacheData cache, SheetInfo d, int col, bool descending, IReadOnlyList<int> rows, CancellationToken token)` that orders only the given rows (stable, same ordering rules). Implement the existing method by calling the new one with all rows from `firstDataRow`.
3. **`CachePolicy.Validate`**: also accept `.xvs` and `.xvs.part` (the style sidecar), nothing else. Add `SidecarPathFor(source)` = cache path with `.xvs`.
4. Run `powershell -File check.ps1`. All old checks must still pass, including `refused ...` checks and `builder refuses outside path`.

Commit: `Prepare the core for the editor: build from an open workbook, sort a row list, allow the style sidecar`.

### Step 2: Host skeleton with WebView2 (2 hours)

1. Create `ExcelApp\ExcelApp.csproj` copying the properties of `ExcelViewer.csproj` (same `NoWarn`, `RollForward`, `SatelliteResourceLanguages`, `DefaultItemExcludes` for `app\**` and `ui\**`), add `<ProjectReference Include="..\ExcelViewer\ExcelViewer.csproj" />` and `<PackageReference Include="Microsoft.Web.WebView2" Version="1.0.2903.40" />`. Set `<ApplicationHighDpiMode>PerMonitorV2</ApplicationHighDpiMode>` and an `<ApplicationIcon>` only if you draw an icon as an `.ico` yourself (a green sheet glyph); otherwise skip the icon.
2. `Program.cs`: `[STAThread] Main`; modes as in 2.2; `Application.SetUnhandledExceptionMode(CatchException)` with the same no-box rule in unattended modes as the old viewer; `--build` calls `CacheBuilder.Build` with the same stdout lines as `ExcelViewer.Program.BuildMode`.
3. `MainForm.cs`: `Text = "ExcelApp"`, 1400 x 860, min 900 x 600, `WebView2` docked `Fill`. Initialise with `CoreWebView2Environment.CreateAsync(null, Path.Combine(Environment.GetFolderPath(LocalApplicationData), "ExcelApp", "WebView2"))`. Settings: `AreDefaultContextMenusEnabled=false`, `AreBrowserAcceleratorKeysEnabled=false`, `IsZoomControlEnabled=false`, `IsStatusBarEnabled=false`, `IsSwipeNavigationEnabled=false`, `AreDevToolsEnabled` only in `--dev`. Map the UI folder: `SetVirtualHostNameToFolderMapping("app.excel", Path.Combine(AppContext.BaseDirectory, "ui"), Allow)` and navigate to `https://app.excel/index.html`; in `--dev` navigate to the given URL. Handle `NewWindowRequested` by cancelling it. Drag and drop of a file onto the form opens it (`AllowDrop`; WebView2 swallows drops, so also handle `CoreWebView2` drop through the UI: the UI sends `openPath` with `File.path` is not available in WebView2 for security; keep the WinForms `DragDrop` on the form and set `webView.AllowExternalDrop = false` so the form receives it).
4. `Bridge.cs`: register handlers by name; parse with `System.Text.Json` (`JsonDocument`); reply as in 2.3; any exception becomes `{ok:false, error: "<plain sentence> (<ExceptionType>)"}` with no message text from Excel (COM messages can quote contents).
5. `Session.cs`: `Open(path)`: check `CachePolicy.ThisPcApproved`; if `CacheData.HeaderMatches` load it, else start the `--build` child and relay progress events, then load. Hold `CacheData`, `Dictionary<int, SheetReader>` guarded by a lock, `hasHeader=true` by default, identity row map. Implement `ready`, `pickFile`, `open`, `workbook`, `sheet`, `rows`, `cellDetail`, `recent`. Recent files: `%LOCALAPPDATA%\ExcelApp\recent.json`, 12 paths, most recent first.
6. A temporary `ExcelApp\ui-stub\index.html` (plain JS, no build) that calls `ready`, shows a button to `pickFile` and `open`, and prints the first 20 rows of sheet 0 as a table. Point the mapping at it until step 3 replaces it; delete it in step 3.
7. Build: `dotnet build ExcelApp\ExcelApp.csproj -c Release`. Run `ExcelApp\bin\Release\net8.0-windows\ExcelApp.exe D:\ExcelAppTest\syn-2m.xlsx` and confirm the stub shows rows; close it. If the WebView2 user data folder cannot be created, fall back to `D:\ExcelViewerCache\..\ExcelAppWebView2`? **No**: use `Path.GetTempPath()` instead and log it. Never put WebView2 data in the cache folder.

Checks: the stub shows made-up rows; `--build` child lines match the old viewer's; `check.ps1` still passes.

Commit: `Add the ExcelApp host: WebView2 window, message bridge, session and cache opening`.

### Step 3: UI skeleton: grid, sheet tabs, status bar (3 hours)

1. `cd ExcelApp\ui`; `pnpm init`; `pnpm add react react-dom @glideapps/glide-data-grid @fluentui/react-icons` and the peers glide asks for (currently `lodash`, `marked`, `react-responsive-carousel`; install whatever `pnpm` warns are missing). `pnpm add -D vite @vitejs/plugin-react typescript @types/react @types/react-dom vitest`. Record the exact versions pnpm chose in `HISTORY.md`. Commit `pnpm-lock.yaml`.
2. `vite.config.ts`: `base: './'`, `build.outDir: 'dist'`, `plugins: [react()]`, `server.port: 5173`. `index.html` loads `/src/main.tsx` and has `<div id="root">` and `<div id="portal">` (glide needs the portal). Import `@glideapps/glide-data-grid/dist/index.css`.
3. `bridge.ts`: `call(cmd, args): Promise<any>` with an id counter and a pending map; `on(event, handler)`; when `window.chrome?.webview` is missing (plain browser during `pnpm dev`), use a **mock host** with a made-up 1,000 x 26 sheet so the UI can be developed in Chrome headless without the host.
4. `rowStore.ts` as in 2.6. `theme.ts` as in 2.6. `Grid.tsx` with the row-number column, letters, selection, zoom.
5. `SheetTabs.tsx`, `StatusBar.tsx`, `App.tsx` with an empty state ("Open a file" button and recent list) and the loading state (progress text from `opening`).
6. `build.ps1`: `pnpm install --frozen-lockfile`, `pnpm exec vite build`, remove and copy `dist\*` to `ExcelApp\app\ui\` after `dotnet publish ExcelApp.csproj -c Release -o ExcelApp\app --no-self-contained`. `Start_ExcelApp.bat` starts `app\ExcelApp.exe`, building first if missing (copy the shape of `Start_ExcelViewer.bat`). Add `ExcelApp\ui\node_modules\`, `ExcelApp\ui\dist\` to `.gitignore` (`app/` is already ignored).
7. Delete `ui-stub`.
8. `--selftest` mode, first version: opens the given made-up file, waits for the `opened` event (poll `ExecuteScriptAsync("window.__app.state()")`), then measures with scripts: `window.__app.scrollTo(viewRow)` to 10 random rows and reports the time from the call until the row store has the visible block and a frame was painted (the UI exposes `window.__app.lastPaintMs`), takes `CapturePreviewAsync` PNGs into the given folder (`open.png`), prints `FIRST SCREEN`, `JUMP median/p95/worst`, `MEMORY host MB`, `MEMORY webview MB` (sum of `msedgewebview2.exe` children by parent process id). Refuse file names not starting with `syn`.

Checks: `build.ps1` succeeds; `app\ExcelApp.exe --selftest D:\ExcelAppTest\syn-20m.xlsx screenshots\` prints first screen under 1,500 ms from process start (cache already built), jump median under 50 ms, and the picture shows letters, row numbers and data. Record the numbers.

Commit: `Add the ExcelApp interface: fast grid, sheet tabs, status bar, self test`.

### Step 4: Ribbon, formula bar, name box, File backstage, View tab (3 hours)

Build exactly section 2.5. Buttons that have no function yet are **not rendered** until their step; Cells and Data Tools groups are rendered disabled with the tooltip because Excel users expect their place. Implement `edge` in the host for Ctrl+arrows. Implement zoom, gridlines and headings toggles, freeze columns and the frozen-rows-in-header drawing, Go To (Ctrl+G) and the name box jump. Keyboard handling in `keys.ts` with a single `keydown` listener on the grid wrapper; make sure glide's own key handling does not double-act (set `keybindings` to turn off what we handle).

Checks: in the selftest add `window.__app.run("freeze", 2, 1)` and `("zoom", 150)` and take `view.png`; the name box reads `A1` after Ctrl+Home; Ctrl+End lands on the last row of the sheet. Chrome headless test of the mock-host UI with Playwright is **not** required (not installed); use the selftest.

Commit: `Add the ribbon, formula bar, name box, File page and View tab`.

### Step 5: Filter and sort (3 hours)

1. `FilterEngine.cs`: `Distinct(sheet, col, candidates, search)` and `Apply(sheet)` → recompute `viewToRow`: start from identity minus header rows; AND every column's condition (scan chunk by chunk through `SheetReader.LoadColumn`, parallel over chunks like `SearchEngine`); then sort by applying `SortEngine.Sort(rows list)` for the last key first and the first key last (stable sort gives multi-key order); prepend header rows. Cancel and recompute when a new condition arrives.
2. Host commands `filterValues`, `setFilter`, `clearFilters`, `setSort`, `setHeaderRow`.
3. `FilterMenu.tsx`: Excel's AutoFilter menu, pixel-close: `Sort A to Z`, `Sort Z to A`, `Sort by Color` (disabled), separator, `Clear Filter From "Letter"`, `Filter by Color` (disabled), `Number Filters` or `Text Filters` submenu (chosen by the column's majority kind in the first loaded block, then confirmed by the host's answer which also returns `kind`), a search box, the check list with `(Select All)` and `(Blanks)`, OK and Cancel. Values list is virtualised (10,000 rows). `SortDialog.tsx` with up to 3 keys and "My data has headers".
4. The Data tab and Home → Sort & Filter wire to these. Row numbers turn blue when a filter is active; the status bar shows `n of m records found`.

Checks (in `ExcelApp.Tests`, step 11 adds the project; create it now with these tests and extend later): filter by values keeps exactly the chosen texts; number `between` and text `contains` are right on a made-up sheet; two columns AND; sort after filter keeps the header first; `Apply` on the 1,000,000-row made-up sheet under 1 s for one column; `Distinct` on it under 1 s. In the selftest: `("filter", col, values)` then `filter.png`.

Commit: `Add Excel-style filtering and multi-key sorting`.

### Step 6: Find and Replace, Go To (1.5 hours)

`FindReplace.tsx` modelled on Excel's dialog (tabs Find / Replace, "Find what", "Replace with", Within: Sheet / Workbook, Match case, buttons Find All, Find Next, Replace, Replace All, Close) using `find` events and `replace`. Find All shows the hit list (sheet, cell, value) under the dialog like Excel; clicking a hit selects the cell (switching sheets when needed; a hit's row is a real row, so map real → view through the host command `locate { sheet, row }` that returns the view index or -1 when hidden by a filter, in which case the UI says "This cell is hidden by a filter").

Checks: a known made-up needle is found once; Replace All on a text column changes the count the test expects; `check.ps1` and the new tests pass.

Commit: `Add Find, Replace and Go To`.

### Step 7: Editing: values, undo, redo, copy, cut, paste, fill (3 hours)

`EditSet.cs` and the commands `setCell`, `clearCells`, `undo`, `redo`, `copy`, `cut`, `paste`, `fill` as in 2.3 and 2.4. The title shows `*` after the file name when dirty (Excel shows nothing, but the owner must see unsaved work; keep the star). The close path: `FormClosing` asks the UI to show the Excel prompt ("Want to save your changes to X?" Save / Don't Save / Cancel) through an event `closeRequested`; the UI answers with `closeAnswer { action }`. In unattended modes never prompt.

Checks: typing `12.5%` into a General cell shows `12.50%`; typing `=A1*2` shows the formula text in italics and `cellDetail.raw` is the formula; undo restores the cache value and the display; paste of 3 x 2 tab text makes 6 edits and one undo record; copy puts the shown text on the clipboard with the three flags (reuse the old `ClipboardSink` test idea on made-up text); fill handle copies a block.

Commit: `Add editing with undo, redo, copy, cut, paste and the fill handle`.

### Step 8: Save through Excel (3 hours)

`SaveBack.cs`, run in the `--save` child:

1. Read the JSON job from stdin: `{ source, target, sameFile, edits: [{ sheet: "name", row, col, kind, num?, text?, formula? , style? }], colWidths: [...], freeze: [...] }`. Rows and cols here are **absolute Excel coordinates** (the host adds `FirstRow`/`FirstCol` before sending).
2. Precondition in the host before starting the child: the source still matches the cache header (`CacheData.IsFreshFor`); if not, refuse with "The file changed on disk since you opened it. Use Save As, or read the file again and redo your changes." and offer Save As.
3. Start hidden Excel exactly like `CacheBuilder.Build` does (Visible false, DisplayAlerts false, AutomationSecurity 3, ScreenUpdating false, EnableEvents false; also `xl.Interactive = false`). Check commit under `Machine.StopAt` first.
4. `wb = xl.Workbooks.Open(source, UpdateLinks: 0, ReadOnly: false, Notify: false, Password: "")`. If `wb.ReadOnly` is true, the file is open elsewhere or read-only: close, quit, exit code 4 with `STOPPED readonly`. The host says "The file is open in another program or is read-only. Close it there and save again, or use Save As."; Save As is still allowed (the copy opens fine).
5. For each sheet with edits: `ws = wb.Worksheets[name]`; if `ws.ProtectContents` is true, exit code 6 `STOPPED protected` ("A sheet is protected, so its cells cannot be changed."). Group edits into contiguous rectangles per row run where possible; otherwise write cell by cell: `ws.Cells[r,c].Formula = formula` for formulas, `.Value2 = num` / `text` / `bool`, `ClearContents()` for cleared cells. Text that starts with `=`, `+`, `-` or `@` but is meant as text is written with a leading apostrophe (`"'" + text`) so Excel keeps it as text. Styles: `Interior.Color` (BGR int), `Font.Bold/Italic/Underline/Color`, `HorizontalAlignment` (xlLeft -4131, xlCenter -4108, xlRight -4152), `NumberFormat`, `Borders` (xlEdgeBottom 9, xlEdgeTop 8, xlEdgeLeft 7, xlEdgeRight 10, inside 11, 12; `LineStyle = 1`), `Columns[c].ColumnWidth`, `ws.Activate(); xl.ActiveWindow.FreezePanes = false; SplitRow/SplitColumn; FreezePanes = true`. Report `WRITING nn` percent every 2%.
6. `xl.Calculate()` is not needed (Calculation is automatic) but call `xl.CalculateFull()` only if `xl.Calculation` was manual in the file; do not change the file's calculation setting.
7. Save: same file → `wb.Save()`; Save As → `wb.SaveAs(target, FileFormat)` with 51 for `.xlsx`, 52 `.xlsm`, 50 `.xlsb`, 56 `.xls`; `ConflictResolution: 2` (local wins). The format follows the target extension the user chose in the dialog. Print `SAVED`.
8. **Rebuild the cache from the same open workbook**: `CacheBuilder.BuildFromOpenWorkbook(xl, wb, target, CachePolicy.PathFor(target), report, cancelled, StyleStore.AfterRead)` (the last argument from step 10; pass null until then). If the user saved to a new path, the new cache is for the new path; the host then switches the session to the new path.
9. Close and quit in `finally` as `CacheBuilder` does; print `RESULT code=... seconds=... edits=...`.
10. Host side: `save` shows `SaveProgress.tsx` (Excel-like small dialog: "Saving through Excel... Writing changes 40%", then "Reading the saved file back 72%"; Cancel is enabled only before `OPENED`), then clears `EditSet`, reloads the cache, keeps the current sheet, selection and scroll position, and sends `opened`. Ctrl+S saves; File → Save As asks the native dialog with filter `Excel Workbook (*.xlsx)|*.xlsx|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel Binary Workbook (*.xlsb)|*.xlsb`, default the source extension and folder.
11. Policy test for the new folder: like `Policy()` in the old tests, scan `ExcelApp\*.cs` for `SaveAs(`, `.Save()`, `FileMode.Create`, `File.WriteAll`, `File.Copy`, `File.Move`, `File.Delete`, `HttpClient`, `WebClient`, `Socket`: allowed only in `SaveBack.cs` (Excel's Save and SaveAs), `Log.cs` (its own log file under `%LOCALAPPDATA%\ExcelApp`) and recent-file writing in `Session.cs` (`recent.json` only). Also assert that no file in `ExcelApp` or `ExcelViewer` mentions `SheetJS`, `xlsx` npm, `ClosedXML`, `OpenXml`, `EPPlus`.

Checks in `ExcelApp.Tests` (live, made-up data): make a workbook with Excel COM (3 sheets, `B2 = 10`, `C2 = =B2*2`, a text column, a formatted money column); build its cache; create an `EditSet` with `B2 = 25`, a text change, a formula `=SUM(B2:B3)` in `D2`, a cleared cell, a bold red fill style on `A1`; run `SaveBack` into a **new** file (Save As) and also into the same file (Save); after each, open the saved file in a fresh hidden Excel and verify: `B2` is 25, `C2` is 50 (Excel recalculated), `D2` has the formula and the right value, the cleared cell is empty, `A1` is bold with the fill, the untouched sheets are byte-for-byte equal in values (compare `UsedRange.Value2` arrays), the first 8 bytes of the saved file are `<## NASC` (print only `NASCA mark: yes/no`), the cache after rebuild shows `25` for `B2` and the formula result for `C2`, the source's cache header matches the saved file (fresh), and no `.part` file is left. Also: a read-only source (set the attribute) gives code 4 and the file is unchanged; a protected sheet (`ws.Protect()` in the generator) gives code 6.

Commit: `Save edits back through Excel and read the saved file in the same session`.

### Step 9: Style edits in the interface (2 hours)

Wire Home → Font, Alignment, Number and Borders to `setStyle`; render with `drawCell` (2.6). Number format changes re-render numbers through `NumFmt` (if `NumFmt` rejects the code, show the raw number with `⚠` as the viewer does; the chosen codes are: Number `#,##0.00`, Currency `[$$-409]#,##0.00`, Accounting `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`, Short Date `yyyy-mm-dd`, Long Date `dddd, mmmm d, yyyy`, Time `h:mm:ss AM/PM`, Percentage `0.00%`, Fraction `# ?/?`, Scientific `0.00E+00`, Text `@`; `%` button `0%`; `,` button `#,##0`; increase and decrease decimals edit the current code's decimal places, General becomes `0.0`). Clear Formats removes the style edit. The Excel colour palette: the 10 theme colours with 5 tints each plus the 10 standard colours (take the hex values from Excel's default Office theme; they are public knowledge).

Checks: the selftest applies bold, fill `#FFFF00`, right alignment and `0.00%` to a range, takes `styles.png`, and `cellDetail` reports the style; the save test of step 8 already proves styles reach the file.

Commit: `Add font, fill, border, alignment and number format editing`.

### Step 10: Read cell styles from Excel into the sidecar (2.5 hours, with a measured gate)

Goal: the grid shows the workbook's own fills, bold text, font colours, alignment, hidden rows and columns, and frozen panes, the way Excel does. Method: the proven **null-if-mixed bisection** of `CacheBuilder.ReadRuns`, applied per column to `Interior.Color`, `Font.Color`, `Font.Bold`, `HorizontalAlignment`, and per sheet to `Rows.Hidden` / `Columns.Hidden` (bisection over row ranges and column ranges), plus `wb.Windows[1]` `SplitRow`, `SplitColumn`, `FreezePanes` per sheet after `ws.Activate()`. Each property returns a value when the whole range agrees and `null` (DBNull) when mixed, so a uniform column costs one call and a striped column costs many. Keep a call budget of 200,000 calls or 120 s per workbook (own constants in `StyleStore`); when the budget ends, remaining runs are "unknown" and draw plain.

1. **Measure first** on made-up data: make a styled workbook with Excel COM (20 columns x 200,000 rows: 10 plain columns, 5 with one fill, 3 with a header-only style, 2 zebra-striped). Time the sidecar build. Write the numbers to `docs\DECISION-3-EXCELAPP.md`.
   - Under 20 s on this file: style reading **on by default**.
   - 20 to 90 s: on by default but with the budget halved; the status bar says `Colours partly read` when runs were cut.
   - Over 90 s: **off by default**; File → Info gets a check box `Read colours and fonts next time` stored in `%LOCALAPPDATA%\ExcelApp\settings.json`. Record which outcome happened.
2. `StyleStore.cs`: sidecar format `XVSTYLE1`, written through `CachePolicy.OpenForWrite(SidecarPathFor(source) + ".part")` and `ReplaceWith`; header holds the cache's `SourceLength`/`SourceTicks` so a stale sidecar is ignored; per sheet: per column runs for the four properties, hidden row runs, hidden column runs, frozen rows and cols. `AfterRead(xl, wb, cache)` is the `afterRead` hook of step 1; it is also passed from `SaveBack`.
3. The host merges sidecar styles under the edit overlay into the wire `style` field; hidden rows and columns are **shown** but drawn with a grey row number and a dotted border line (the owner decided in the viewer that hidden rows and columns stay visible; keep that and make it visible that Excel hides them). Frozen panes from the file are applied on open.

Checks: on the styled made-up workbook the sidecar reads back the fills and bold runs the generator wrote (test compares expected run lists); a plain workbook builds a near-empty sidecar in under 1 s on top of the cache build; `check.ps1` still passes (the old viewer ignores sidecars).

Commit: `Read fills, fonts, alignment, hidden rows and frozen panes from Excel into a style sidecar`.

### Step 11: Tests, self test and measurements (2 hours)

1. `ExcelApp.Tests\Program.cs` holds every test named in steps 5 to 10 plus: typed-input parsing table (numbers in the current culture, percent, dates, TRUE, formulas, text with leading `=` typed as text through the apostrophe rule), `EditSet` undo and redo order, the `copy` text with gaps, `paste` splitting with quoted cells (`CopyText.Quote` inverse), `rows` wire format for a row with every kind, the policy scan of step 8, and the memory stop (a fake `Machine.Commit` is not possible; test that `SaveBack` refuses when passed a `stopAt` of 0 through an internal parameter).
2. `check.ps1`: build and run both test projects; `-Offline` passes `offline` to both. Exit code is non-zero if either fails.
3. Full selftest run on `syn-20m.xlsx` with the complete UI: record first screen, jump median/p95/worst, filter apply on one column, sort, find nothing found, memory of the host and the WebView2 processes, and a save round trip of 100 edits on `syn-2m.xlsx` (time from Ctrl+S to `opened`). Pictures into `screenshots\` (gitignored).
4. Acceptance (made-up data, this PC): first screen from cache under 1,500 ms from process start; jump median under 50 ms and worst under 200 ms; filter and sort on a 1,000,000-row column under 1.5 s each; host plus WebView2 memory under 700 MB at 20 million cells; save of 100 edits on 2 million cells under 60 s including the read back. If a number misses, spend at most one hour on it, then record the number honestly and continue.

Commit: `Add the ExcelApp tests and measurements`.

### Step 12: Documents and clean-up (1 hour)

1. `README.md`: a new section **ExcelApp** before "The rule": what it is (the editor for very large files, looking like Excel), how to start (`ExcelApp\Start_ExcelApp.bat`), what it does (open, scroll, filter, sort, find, edit, save through Excel so files stay protected), what it does not do yet (charts, pivots, macros, recalculating formulas inside the app, inserting rows and columns, merged cells drawn as one), and that the first opening of a file still waits for Excel. Update the Layout table with `ExcelApp`, `ExcelApp.Tests`, `docs\PLAN-3-EXCELAPP.md`, `docs\DECISION-3-EXCELAPP.md`.
2. `HISTORY.md`: a 2026-11-07 (or the actual date) section listing each step's commit, the measured numbers, every deviation from this plan and why, package versions chosen, and the style-reading gate outcome.
3. `docs\DECISION-3-EXCELAPP.md`: the measurements tables (made-up data), the style gate result, the save round-trip numbers, and open points for phase 2.
4. `.gitignore`: add `Browser Excel/` with the comment `# The owner's browser experiments (Univer, glide-data-grid): not part of the product`, `ExcelApp/ui/node_modules/`, `ExcelApp/ui/dist/`, `*.log`, `D:/ExcelAppTest` is outside the repo so nothing needed.
5. Run `git status` and make sure no `.xvc`, `.xvs`, `.png` of real data, `node_modules`, `dist`, `app` or `Browser Excel` content is staged.

Commit: `Document ExcelApp and ignore the browser experiments`.

---

## 4. Things that are easy to forget (all required unless marked later)

- **File already open in the user's own Excel**: reading works (we open read-only); saving is refused with the message in step 8.4, and Save As still works.
- **Read-only files and files on network shares**: open works; the cache is still local; Save is refused, Save As works. The cache folder rule is unchanged: nothing about the network share is written anywhere except `recent.json` (the path only).
- **Password-protected workbooks**: `Open` with `Password: ""` throws; the child exits code 1 and the host says "This file needs a password, which is not supported." No prompt can appear because `DisplayAlerts=false` and `Interactive=false`.
- **xlsm**: macros are never run (`AutomationSecurity=3`); saving keeps them (format 52). **xlsb and xls**: open and save keep their format. **CSV**: open works through Excel; saving a CSV is refused ("Save this as an Excel workbook instead") because Excel's CSV save drops sheets and prompts.
- **The file changed on disk while open**: a `FileSystemWatcher` on the source sends `fileChanged`; the UI shows a bar "This file changed on disk. Read it again to see the changes." Save is then refused (step 8.2).
- **Very wide sheets** (up to 16,384 columns): letters up to `XFD`; `BlockRows` already shrinks for wide sheets; the UI must not create one React element per column (glide does not).
- **Empty sheets and workbooks with only hidden sheets**: show the first visible sheet or an "This sheet is empty" state; never crash on `Rows == 0`.
- **Merged cells**: drawn as separate cells with the value in the top-left cell (that is what the cache holds). Say so in README. Later: read `MergeCells` runs.
- **Hidden rows and columns**: shown, marked grey (step 10). Later: a View toggle to hide them.
- **Formulas**: the cache holds values only. Edited formulas are written to Excel and recalculated on save; the cache is rebuilt from the saved file so results appear after save. Reading formula text into the cache is phase 2 (one more `Range.Formula` block read per block and a `XVCACHE5` magic).
- **Dates typed by the user**: see 2.4; the number format applied makes the date visible. Dates read from the file already show as Excel shows them.
- **Culture**: the UI shows numbers as the cache renders them (Excel's own formats); typed input is parsed in the Windows culture of the user. Never hard-code `.` as the decimal separator in parsing.
- **High DPI and multiple monitors**: `PerMonitorV2`; glide uses `devicePixelRatio`; test the selftest picture is sharp (no blurry text) by checking the canvas backing size equals CSS size times DPR.
- **Memory**: `memory` events above 85% show a status bar warning; the host refuses to start a build or save above `Machine.StopAt`; the row store LRU keeps at most 400 blocks of 200 rows.
- **Crash isolation**: Excel runs only in children. If a child dies, the host shows one plain sentence and stays alive. The host's unhandled exception path never shows cell text.
- **Clipboard history**: every copy uses the three flags. Cut leaves the clipboard content; the source is cleared on the next paste inside the app only.
- **Closing with unsaved changes**: the Excel-style prompt, step 7. Closing during a save waits for the child to finish (it owns an Excel) with a "Finishing the save..." state.
- **Second instance**: allowed; each has its own session. Two instances saving the same file at once: the second gets the read-only refusal.
- **Recent list and settings**: paths only, in `%LOCALAPPDATA%\ExcelApp`. Never a cell value.
- **Logging**: `Log.cs` writes timings and counts; the `log` command rejects long strings; the log never receives a path of a real file (hash it like `CachePolicy.PathFor` does).
- **.NET 8 end of support on 2026-11-10**: build with 8, run on 10 through `RollForward=LatestMajor` (already how the viewer works). Note it in README.
- **The WebView2 runtime** is installed by Windows 11 (Edge). If `CoreWebView2Environment.CreateAsync` throws, the host shows one message naming the WebView2 runtime and exits with code 7. Do not download anything.
- **Phase 2 (not in this run)**: insert and delete rows and columns, merged cells, hide rows and columns, formula text in the cache and a formula bar that shows file formulas, conditional formatting display, data validation lists, comments and notes, freeze rows if the header trick failed, column autofit from content, Excel's `Remove Duplicates` and `Text to Columns`, dark mode, printing.

---

## 5. What "done" looks like

- `powershell -File check.ps1` passes with both test projects, including the live Excel tests on made-up workbooks.
- `ExcelApp\Start_ExcelApp.bat` opens the window; `syn-20m.xlsx` (cache already built) shows its first screen in under 1.5 s; scrolling, jumping, filtering, sorting, finding, editing and saving all work in that window.
- `app\ExcelApp.exe --selftest D:\ExcelAppTest\syn-20m.xlsx screenshots\` prints the acceptance numbers and writes `open.png`, `view.png`, `filter.png`, `styles.png`.
- The saved file after a round trip starts with the NASCA mark, opens in Excel with the edits and recalculated formulas, and the app shows the saved state without a second Excel wait.
- `git log` shows the eleven commits named above on `main`; `git status` is clean except the owner's `Browser Excel\` folder, now ignored.
- `HISTORY.md`, `README.md` and `docs\DECISION-3-EXCELAPP.md` are updated; no real data anywhere in the repo.

## 6. The final report (write it as the last message)

Ten to twenty lines, plain English: what was built, the measured numbers (made-up data), every deviation from this plan, the style-reading gate outcome, anything that did not pass and why, and the three things the owner should try first. No cell values, no real file names.
