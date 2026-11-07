# Plan 4: the data engine — auto-ingestion and the 20 ideas, one shot

Source brief: the owner's `ExcelApp_Developer_Prompt_20_Ideas.md` (2026-11-07). The vision: ExcelApp becomes **a local database with a spreadsheet face**. The `.xlsx` file is the NASCA-approved packaging, and the work happens in the fast store. People ask questions and get answers instead of opening files and reading rows.

This plan is complete. Three coders build it at the same time in three lanes (A, B, C), each in its own git worktree. Every lane reads **this whole file** first. Then a reviewer reviews and fixes the merged result, and the lead checks and finishes it. A coder never guesses: if this plan does not say how, choose the simplest way that keeps every rule in section 0, write the choice in your answer file and continue. **Do not stop to ask. Finish every step in your lane.**

## 0. Rules that never bend (all lanes)

1. **Real data stays on this laptop.** Cell data is written only through `StoreRoot` (`ExcelApp/Data/StoreRoot.cs`). On the approved laptop that is `D:\ExcelViewerCache\store\`, checked by `CachePolicy.ValidateStore`. Configuration goes only through `ConfigRoot` (`%APPDATA%\ExcelApp\config`) and holds no cell data, apart from rule values the user typed. **Never** write data with `File.*` directly, to `%TEMP%`, or to any other folder.
2. **Excel is the only reader and writer of company files.** Every read of an `.xlsx`/`.xls`/`.csv` goes through the existing hidden-Excel child (`ExcelApp.exe --build <file>`). Every export goes through the new hidden-Excel child (`ExcelApp.exe --export <job>`). Never open a company file with a ZIP/XML/CSV reader, never pass its bytes to a parser, and never strip or avoid NASCA. Hashing a file's bytes (SHA-256, for identity only) is allowed.
3. **Never kill Excel and never touch the owner's Excel.** A child owns its own hidden instance, and closes only what it opened.
4. **No new packages** (NuGet or npm). Section 1 explains why. Use .NET 8 base libraries, the existing projects and the existing UI dependencies. Charts are hand-drawn SVG.
5. **Never log** cell values, sheet names, column names from real files, or real file names to `Log.cs`, the console or the answer files. The app's own screens may show them.
6. **Made-up data only in tests.** Test folders go under `D:\ExcelAppTest\syn*` (`StoreRoot.ForTest`, `ConfigRoot.ForTest`). Test workbooks are created by Excel through COM from made-up values. Never open `D:\ExcelViewerCache` or any `.xvc`.
7. **No Windows error box, ever.** Programs you run (tests, children, throwaway consoles) start with `SetErrorMode(0x8003)`, or go through `python ExcelApp/run-guarded.py <command>`.
8. **No git commands, no other agents, no agent_bridge.** The lead commits and merges. Codex lanes are Herdr delegates: do the work yourselves.
9. **Every feature works on millions of rows without the interface waiting.** Anything over 200 ms runs off the UI thread and reports progress through an event.
10. **Deterministic rules first.** There is no AI model in this build (section 7).
11. **Own only your files** (section 3). Need a change in a lead-owned contract file? Write the request in your answer file, and add a clearly marked local adapter in your own files meanwhile.

## 1. Environment facts that shaped this plan (measured 2026-11-07, 21:20)

- Windows 11 Enterprise, 8 logical CPUs, 15.7 GB RAM, commit limit 18.9 GB. Free disk: **C: 3.5 GB, D: 13.9 GB**. Build outputs stay on D:.
- Tools: .NET SDK **8.0.427 only** (no .NET 10 SDK), PowerShell 7.6.9, Node 25.2.1 with pnpm 11.7.0, Python 3.12.10, Codex CLI 0.166.0, Excel 16 x64.
- Network: nuget.org answers through the proxy, but the DuckDB native package (110 MB) stalled at 8 MB. GitHub release downloads time out. **DuckDB is therefore not used.** The fast store is our own columnar format in plain C# (section 2.2), behind `ITableProvider`, so DuckDB can be swapped in later if a network allows it.
- .NET 10: not installed, and C: is nearly full. The projects stay on `net8.0-windows` with `RollForward=LatestMajor` (already set), so the app runs on a .NET 10 runtime once one is installed. The SDK move is an owner decision (section 7).

## 2. The design, fixed

### 2.1 Process model
- `ExcelApp.exe` (the host) runs the window, the editor `Session`, and the **DataHost** (`ExcelApp/Data/DataHost.cs`). DataHost creates `DataServices` and calls `DataCommands.Register` (lane A), then `QueryCommands.Register` (lane B), then `Ingestion.Start()`.
- **One background worker queue** (lane A, `WorkQueue`) runs every job that needs Excel: ingestion reads, workspace reads, ingestion previews and exports. Only one hidden Excel child runs at a time (`Settings.MaxParallelExcel = 1`).
- Children:
  - `ExcelApp.exe --build <file>`: exists. It writes the `.xvc` fast copy through CachePolicy and prints `RESULT code=...`.
  - `ExcelApp.exe --export <job path>`: new, lane A. It writes an `.xlsx` through Excel.
- The interface talks to the host with the existing bridge: `call("data.xxx", args)` and `on("data.yyy")`. Commands and events are listed in section 4.
- The engine runs while ExcelApp is open. `Settings.StartWithWindows` adds `HKCU\Software\Microsoft\Windows\CurrentVersion\Run\ExcelApp = "<exe> --tray"`. It is off by default, and the user turns it on in Settings. `--tray` starts the window minimised.

### 2.2 The table file `.xvt` (lane A writes it, everyone reads it through `ITableProvider`)
Little-endian. Each section starts on an 8-byte boundary.
```
magic   "XVTABLE1" (8 bytes)
int32   formatVersion = 1
int32   headerLength
byte[]  header JSON (UTF-8): { "name": str, "rows": int, "columns": [ { "name": str, "type": "number|text|date|bool",
         "offset": long, "length": long, "hasNulls": bool, "dictCount": int } ], "meta": { free-form string map } }
sections, one per column, at "offset":
  number/date: double[rows]; then, when hasNulls, byte[rows] (1 = null)
  text:        int32[rows] codes (-1 = null); then int32 dictCount; then per entry int32 byteLength + UTF-8 bytes
  bool:        sbyte[rows] (0, 1, -1 = null)
```
- `TableFile.Write(FileRoot root, string relative, Table t, IReadOnlyDictionary<string,string> meta)` writes `relative.part` and then commits.
- `TableFile.Open(FileRoot root, string relative)` reads only the header. Columns load lazily on first use, through `MemoryMarshal.Cast<byte,double>` on one read per section. The result is cached as a `Table` whose columns are loaded on demand. Use a small `LazyTable` wrapper inside lane A, and give `Table` to callers once loaded.
- Writing 10 million rows by 10 columns must take under 10 seconds. Opening its header must take under 50 ms.

### 2.3 Store layout (under `StoreRoot`; lane A owns all of it, except the folders marked B)
```
catalog.json                           report types, projects, tables, the next global version id
reports/<rt>/versions.json             VersionsFile: periods -> versions, current, locked, lockedVersion
reports/<rt>/<period>/v<N>.xvt         typed, cleaned rows of period version N (with _v, _r)
reports/<rt>/<period>/v<N>.raw.xvt     the same rows before cleanup, all Text (makes cleanup reversible)
reports/<rt>/<period>/v<N>.diff.xvt    compare result against the previous current version (_status column)
tables/<project>/<table>.xvt           workspace tables (idea 2), with their raw twin
ingest/log.jsonl                       one LoadRecord per line (newest last)
ingest/seen.json                       file hash -> load id (duplicate guard)
ingest/quarantine/<loadId>.xvt + .json staged data waiting for the user
exports/<jobId>.json                   export job (holds data; deleted by the child when done)
alerts/log.jsonl                       (B) fired alerts
sandboxes/<id>.json + <id>.xvt         (B) sandbox edits and edited columns
journal/                               (lead) editor auto-save journal
```
Configuration (under `ConfigRoot`): `recipes/<id>.json` (ingest recipes, lane A), `views/`, `links/`, `alerts/`, `reportRecipes/`, `dashboards/`, `projects.json` (lane B), and `settings.json` (lane A). Every save is atomic, through `WriteAllText`.

### 2.4 The data model: report type → period → version
- A **report type** is one ingestion recipe (`IngestRecipe.Id`). A **period key** is `yyyy-MM-dd` for a day, `yyyy-Www` for an ISO week (`FYyyyy-Www` for a fiscal week, see WeekRule), and `yyyy-MM` for a month.
- `VersionsFile` (lane A type, JSON):
  `{ "nextVersion": int, "periods": { "<period>": { "versions": [VersionInfo...], "current": N, "locked": bool, "lockedVersion": N|null } } }`.
  `VersionInfo` matches `contracts.ts`, plus `versionId` (the global id in `_v`) and `status` (`current` | `archived` | `rolledBack`).
- **Latest truth**: for each period use `current`. **Locked truth**: use `lockedVersion` for locked periods and `current` for the others. Versions loaded after the lock are stored with `postClose = true`.
- **As of T**: for each period, use the highest version with `loadedAt <= T` that is not rolled back. **As of a version**: `TableRef.Version` selects a period version, and `kind "version"` selects a global version.
- **The current report table** is the union of the chosen version of every period, in period order. Union is by column name; a column missing in one period gives nulls there. It is cached until the next load of that report type.
- **Lock rule**: a period locks at 23:59:59 of the `WorkingDay`-th working day (Monday to Friday, minus `Holidays`) of the next period. It is checked at Start and every hour. Locking sets `locked = true` and `lockedVersion = current`. The user can also lock or unlock by hand (`data.history.lock`).
- Every version keeps `_v` (the global version id) and `_r` (the 1-based source sheet row). `SourceOf(versionId)` answers file, sheet, header row, load time and column map. A cell address is `sheet!<column letter of SourceColumn[col]><_r>`.

### 2.5 Ingestion pipeline (lane A; plan brief section 2 is binding)
Every file goes through five steps. Each step emits `data.ingest` with its stage, and every outcome appends a `LoadRecord` to `ingest/log.jsonl`.
1. **Notice**
   - A `FileSystemWatcher` runs on each folder of each enabled recipe (Created, Changed, Renamed, with `IncludeSubdirectories = false`). A catch-up scan runs at Start, then every 5 minutes, and on `data.ingest.scanNow`.
   - Candidates are files that match a recipe pattern (a glob: `*` and `?` only, case-insensitive). Ignore `~$*`, `*.tmp`, `*.crdownload`, `*.partial` and `*.part`.
   - Debounce: size and last-write time must be unchanged for 3 checks 2 s apart, and the file must open with `FileShare.Read` (no writer holds it). Then queue it once. The same path and timestamp are never queued twice.
2. **Recognize**
   - Read the file through the Excel child into the fast copy, then read the recipe's sheet with `SheetReader`. The sheet is the recipe's `Sheet`, or the first visible sheet.
   - Header row: the recipe's `HeaderRow`. When it is 0, use the row among the first 20 that matches most `HeaderSignature` names. Compare names normalised: trim, lower case, collapse inner spaces.
   - Signature score = matched names ÷ expected names. Pattern match with score ≥ 0.8 means recognised. With several recipes, take the highest score.
   - Pattern matches but score < 0.8: quarantine, reason **"The columns do not match the recipe"**. List what is missing and what is new.
   - Period: compute it from the `DateColumn` values **inside** the file (the file name is never used). Rows with a blank or unreadable date are counted in a `badDate` flag and excluded.
3. **Compare**
   - Type and clean the rows (2.6), giving a staging table.
   - For each period, compare against that period's current version by `KeyColumns`. Duplicate keys are paired by occurrence order. "Changed" means any non-reserved column's text differs.
   - The diff table has the key columns, `_status`, and for changed rows the old and new value of each changed column, as `<col>` and `<col> (before)`.
4. **Update**
   - Run the safety rules (2.7). If any stops the load, the staging data goes to `ingest/quarantine/` and the status is `quarantined`, with the reason in a plain sentence.
   - Otherwise, for each period, write `v<N>.xvt`, `v<N>.raw.xvt` and `v<N>.diff.xvt`. Then update `versions.json` atomically (one writer lock per report type), set `current = N` and archive the old version.
   - Record `Versions[period] = N`, and the `Added`/`Removed`/`Changed` totals.
5. **Notify and refresh**
   - Raise `IIngestion.Loaded(record)` and emit `data.loaded`.
   - Emit `data.notice` with one plain sentence, built per period: `"October updated (312 changed, 40 added), November added."` A period is "added" when it is new, and "updated" when it had versions.
   - Lane B listens to `Loaded`, so views, links, alerts, dashboards and recipes refresh.

### 2.6 Typing and cleanup (idea 4; lane A, `Typing.cs`)
For each column of the raw sheet (all cells read as shown text, kind and number from the fast copy):
- **Recipe fixes first** (`TypeFix`):
  - `date`: parse text with `DateFormat`, or detect from `dd.MM.yyyy`, `dd/MM/yyyy`, `MM/dd/yyyy`, `yyyy-MM-dd`, `yyyyMMdd` and `dd-MMM-yyyy`, choosing the first format that parses ≥ 95% of the non-empty cells. Numbers that carry a date format are dates.
  - `number`: invariant culture, then the current culture; strip thousands separators, a trailing `%` (÷100), a currency symbol and spaces.
  - `text`: keep as text.
  - `padZeros(width)`: left-pad digit-only text to `width`.
  - `trim`, `upper`.
- **Automatic, when no recipe fix exists:**
  - **Number** when ≥ 95% of the non-empty cells are numbers or numeric text. Numeric text converted is reported as `numberFromText`.
  - **Date** when the cache format of ≥ 95% of the numbers is a date format (`NumFmt` date detection), or the text parses as dates under one format (`dateFromText`).
  - **Text** otherwise.
  - A digit-only text column where any value has a leading zero is **always Text**: codes never lose their zeros. Report it as `leadingZerosKept`.
  - Trailing and leading spaces are trimmed (`trimmed`).
  - Mixed case is only reported (`mixedCase`), never changed, unless the recipe says `upper`.
- Every fix adds a `CleanupItem(column, fix, cells, example)`. `example` is shown in the app only, never logged.
- **Reversible**: `data.ingest.undoCleanup {loadId, column, fix}` writes a new period version re-derived from `v<N>.raw.xvt` without that fix. It also removes the fix from the recipe, and asks the user first in the interface.

### 2.7 Safety rules (never replace blindly)
| Rule | Test | Action |
|---|---|---|
| Duplicate download | Same `FileHash` already in `seen.json` for this recipe, **or** same `ContentHash` as the period's current version | Status `duplicate`, log only, no notice |
| Partial file | New period rows < old rows × (1 − `PartialThreshold`), or an expected signature column is missing | Quarantine: "This file has far fewer rows than the version it would replace (N against M)." |
| Schema drift | New, renamed or missing columns compared with `HeaderSignature` | Quarantine; accept offers "update the recipe" (`updateRecipe: true` adds or removes signature names) |
| Two periods | The dates span more than one period of the granularity | Quarantine; accept offers `split: true`, which loads each period as its own version |
| Future dates | Any date later than today + 1 day | Quarantine |
| Rollback | `data.ingest.rollback {loadId}` | Each period's `current` becomes the previous non-rolled-back version. The load status becomes `rolledBack`, and `Loaded` is raised again |

`accept` and `reject` act on a quarantined load. Accept continues at step 4 without that guard. Reject deletes the staging files and keeps the log entry.

### 2.8 Quality card (idea 16; lane A, `Quality.cs`)
Run after every load, on the new period data. Flag, never block (blocking is only 2.7):
- `duplicateKey`: key used more than once.
- `blank`: blanks in key or date columns, or in a column that was fully filled in the previous version.
- `badDate`: unreadable dates.
- `outOfRange`: negative values in a column that had none in the last 6 versions.
- `outlier`: a numeric value ≥ 10× the median absolute value of that column in the previous version, or a per-key change ≥ 10×.
- `newCategory`: a text column with ≤ 200 distinct values gains a value never seen in the last 6 versions.
- `futureDate`, `twoPeriods`: echoed from 2.7 when accepted anyway.

Each flag keeps up to 20 sample rows. `data.quality.rows` turns a flag into a result that the interface pages, so the flagged rows are one click away.

### 2.9 Workspace (idea 2; lane A)
- A **project** is a named list of files. `data.workspace.add` queues an Excel read of a file. Every visible sheet with at least one non-empty row becomes a table `tables/<project>/<table>.xvt`, typed with 2.6 (automatic mode). It is named `"<file stem> / <sheet>"` and gets a global version id.
- `refresh` re-reads files whose length or last-write time changed. While a file reads, its old table stays usable, marked `refreshing` in `TableInfo.Group`. Opening a project never waits.

### 2.10 Export through Excel (ideas 18 and 19; lane A child, lane B jobs)
- `Exporter.ExportAsync(job)` writes `exports/<id>.json` and runs `ExcelApp.exe --export <full path>` hidden, with a 10-minute timeout. It reads the child's stdout: `PROGRESS <text>`, then `RESULT ok` or `RESULT error <plain sentence>`. A child that times out is left alone (never killed); the export reports "Excel is still busy" and the job file is cleaned up at the next start.
- The child, following `SaveBack.cs`'s COM pattern:
  1. Start its own hidden Excel with `DisplayAlerts = false`, `EnableEvents = false`, `AutomationSecurity = 3`.
  2. If there is a `Template`, open it read-only. Otherwise make a new workbook.
  3. For each `ExportSheet`, find or create the sheet. Resolve `StartCell` as an A1 address or a defined name. Write headers (bold only without a template), then the rows as one `object[,]` through `Value2`: numbers as double, dates as OADate with the column's `NumberFormats`, text as text. Auto-fit only without a template.
  4. `SaveAs(Target, 51)`. If `Pdf` is set and the owner allowed it, `ExportAsFixedFormat(0, <target>.pdf)`.
  5. Close only its own workbook and app. Delete the job file, and print the result.

## 3. Lanes, files and ownership

| Lane | Agent | Owns (creates or edits) |
|---|---|---|
| **A: store and ingestion** | Codex GPT-6.1-Terra medium | `ExcelApp/Data/DataCommands.cs`, `ExcelApp/Data/ExportChild.cs`, new `ExcelApp/Data/Store/**` (TableFile, Catalog, Versions, TableProvider, WorkQueue, ExcelReader, Typing, Quality, Exporter, Settings store), new `ExcelApp/Data/Ingest/**` (Recipes, Watcher, Recognizer, Periods, Compare, Ingestion, History), `ExcelApp.Tests/StoreTests.cs`, `ExcelApp.Tests/IngestTests.cs` |
| **B: analysis, automation, formulas** | Codex GPT-6.1-Terra medium | `ExcelApp/Data/QueryCommands.cs`, new `ExcelApp/Data/Query/**` (QuerySpec, Expr, Engine, Join, Fuzzy, Format, Views, Links, Drill, Alerts, ReportRecipes, Scheduler, Sandbox, Library, Dashboards, Questions), new `ExcelApp/Formula/**`, `ExcelApp.Tests/QueryTests.cs`, `ExcelApp.Tests/AutomationTests.cs`, `ExcelApp.Tests/FormulaTests.cs` |
| **C: interface** | Codex GPT-6.1-Terra medium | new `ExcelApp/ui/src/data/**` **except** `contracts.ts`; `ExcelApp/ui/src/App.tsx` **only** to add the mode switch (step C1); `ExcelApp/ui/src/mockHost.ts` **only** to forward `data.*` to `data/mockData.ts`; `ExcelApp/ui/src/styles/data.css` (new) |
| **Lead** | Claude | `Contracts.cs`, `StoreRoot.cs`, `DataHost.cs`, `contracts.ts`, `CachePolicy.cs`, `Session.cs`, `MainForm.cs`, `Program.cs`, `Bridge.cs`, `EditSet.cs`, `ExcelApp.Tests/Program.cs`, `check.ps1`, docs, editor auto-save, formula integration into the editor |

Everything not listed is read-only for every lane. Lanes read the existing code they need: `Session.cs` for how the build child is started and read, `SaveBack.cs` for COM, `CacheFile.cs` for `SheetReader`, `NumFmt.cs` for dates, `Grid.tsx` and `rowStore.ts` for the grid and paging, and `bridge.ts`.

## 4. Wire protocol: every `data.*` command (args → result)
Errors are thrown as `UserError("plain sentence")`, which the bridge turns into a rejected promise. JSON is camelCase. Types are in `contracts.ts`.

**Lane A**
| Command | Args | Result |
|---|---|---|
| `data.tables.list` | – | `{tables: TableInfo[]}` |
| `data.columns` | `{source: TableRef}` | `{columns: ColumnInfo[]}` (no `_` columns) |
| `data.workspace.projects` | – | `{projects: [{id, name, tables: TableInfo[]}]}` |
| `data.workspace.create` / `.rename` / `.delete` | `{name}` / `{id, name}` / `{id}` | `{id}` / `{}` / `{}` |
| `data.workspace.add` | `{projectId, path}` | `{queued: true}`; then `data.notice` |
| `data.workspace.remove` / `.refresh` | `{projectId, tableId}` / `{projectId}` | `{}` |
| `data.ingest.recipes` | – | `{recipes: IngestRecipe[]}` |
| `data.ingest.saveRecipe` / `.deleteRecipe` | `{recipe}` / `{id}` | `{recipe}` (validated, id slugged) / `{}` |
| `data.ingest.preview` | `{path}` | `{sheets: [{name, rows, cols, headerRow, header: string[], sample: string[][] (20 rows), guesses: [{name, type, fixes: string[]}]}]}` (runs the Excel child; emits `data.ingest` progress) |
| `data.ingest.status` | – | `{running, watching: [{folder, ok, error?}], queue, current?}` |
| `data.ingest.scanNow` | – | `{}` |
| `data.ingest.loadFile` | `{path, recipeId}` | `{queued: true}` (manual load, same pipeline) |
| `data.ingest.log` | `{limit?, recipeId?}` | `{records: LoadRecord[]}` newest first |
| `data.ingest.quarantine` | – | `{records}` |
| `data.ingest.accept` | `{loadId, split?, updateRecipe?}` | `{record}` |
| `data.ingest.reject` / `.rollback` | `{loadId}` | `{record}` |
| `data.ingest.undoCleanup` | `{loadId, column, fix}` | `{record}` |
| `data.history.versions` | `{reportTypeId, period?}` | `{periods: string[], versions: VersionInfo[]}` |
| `data.history.diff` | `{reportTypeId, period, from, to}` | `DiffResult` (its `resultId` pages with `data.page`) |
| `data.history.trend` | `{reportTypeId, column, agg: "sum"|"count"|"avg", mode: "versions"|"periods", period?, last?}` | `{points: [{label, value, version?, loadedAt?}]}` |
| `data.history.lock` | `{reportTypeId, period, locked}` | `{}` |
| `data.quality.rows` | `{loadId, flag}` (index) | `ResultInfo` |
| `data.settings.get` / `.set` | – / `{settings}` | `{settings}` |

**Lane B**
| Command | Args | Result |
|---|---|---|
| `data.run` | `{spec: QuerySpec}` | `ResultInfo` |
| `data.page` | `{resultId, from, count ≤ 1000}` | `ResultPage` (works for every `ResultStore` entry, including A's diff and quality results) |
| `data.release` | `{resultId}` | `{}` |
| `data.drill` | `{resultId, row}` | `ResultInfo` (the rows behind a grouped row) |
| `data.lineage` | `{resultId, row, col}` | `LineageCell` (ungrouped results) |
| `data.export` | `{resultId, target, template?, sheet?, startCell?, pdf?}` | `{ok: true}` |
| `data.views.list` / `.save` / `.delete` / `.run` | – / `{view}` / `{id}` / `{id}` | `{views: (SavedView & {count, refreshedAt})[]}` / `{view}` / `{}` / `ResultInfo` |
| `data.links.list` / `.save` / `.delete` / `.run` | same pattern | `{links}` / `{link}` / `{}` / `ResultInfo` |
| `data.alerts.list` / `.save` / `.delete` | same pattern | `{alerts}` / `{alert}` / `{}` |
| `data.alerts.fired` | `{limit?}` | `{events: AlertEvent[]}` |
| `data.recipes.list` / `.save` / `.delete` | same pattern | `{recipes: ReportRecipe[] & lastRun}` |
| `data.recipes.run` | `{id}` | `{ok, message, resultId?}` |
| `data.sandbox.create` | `{base: TableRef, name}` | `Sandbox` |
| `data.sandbox.list` / `.discard` | – / `{id}` | `{sandboxes}` / `{}` |
| `data.sandbox.edit` | `{id, edits: [{row, col, value}]}` | `{edits}` |
| `data.sandbox.compare` | `{id}` | `DiffResult` |
| `data.sandbox.keep` | `{id, reportTypeId, period, note}` | `{version}` |
| `data.library.preview` | `{items: [{kind, id}]}` | `{literals: [{kind, id, field, value}]}` |
| `data.library.export` | `{items, target}` | `{ok}` |
| `data.library.import` | `{path}` | `{imported: [{kind, id, name}], renamed: string[]}` |
| `data.dashboard.list` / `.save` / `.delete` | same pattern | `{dashboards}` / `{dashboard}` / `{}` |
| `data.dashboard.values` | `{id}` | `{values: TileValue[]}` |
| `data.ask` | `{text}` | `Question` |

**Events** (host → interface): `data.ingest`, `data.loaded`, `data.quarantine`, `data.alert`, `data.recipeRun`, `data.notice`. Their shapes are in `contracts.ts`.

## 5. The steps

### Lane A: store and ingestion (Codex GPT-6.1-Terra medium, worktree `Excel-wt-a`, branch `plan4-a`)
Work in this order. After each step, run `powershell -File check.ps1 -Offline` (all green) and append progress to the answer file.

- **A1. Table files (2.2).** Write `Store/TableFile.cs` (Write, Open, LazyTable) and `Store/Catalog.cs` (catalog.json: report types, projects, tables, next global version id; with a lock).
  - Tests in `StoreTests`: a round trip of every type with nulls, an empty table, a 1-row table, Unicode text, a 2 GB guard (refuse a table over 2³¹ rows), and a half-written `.part` that is never read.
  - Speed: 10 M rows by 10 columns written in < 10 s; the header opens in < 50 ms.
- **A2. Excel reader and work queue.**
  - `Store/ExcelReader.cs` runs `ExcelApp.exe --build <file>` as `Session.cs` does: same arguments, output parsing, hidden window and timeout. Then it loads the `.xvc` with `CacheData.Load(CachePolicy.PathFor(file))`, and offers `RawSheet ReadSheet(string file, string? sheet, CancellationToken)`.
  - `RawSheet` has `Rows`, `Cols`, `Kind(r, c)`, `Num(r, c)`, `Text(r, c)` (shown text) and `IsDateFormat(r, c)`.
  - Behind an `IRawSource` interface, so tests feed in-memory sheets.
  - `Store/WorkQueue.cs`: one worker thread, `Enqueue(string label, Func<CancellationToken, Task>)`, and status.
- **A3. Typing and cleanup (2.6)** in `Store/Typing.cs`: `TypedResult Type(RawSheet s, int headerRow, IReadOnlyList<TypeFix> fixes)` returns a `Table` (with `_r`) plus a `CleanupReport` plus the raw `Table`. Test every rule, including the leading-zeros trap, `%` and thousands separators, and 6 date formats.
- **A4. Periods** in `Ingest/Periods.cs`:
  - `string PeriodOf(double serial, Granularity g, WeekRule w)`.
  - `DateTime LockTime(string period, Granularity g, LockRule r)`.
  - Tests: ISO week 53, year boundaries, a fiscal year start, holidays and weekends in the lock rule.
- **A5. Versions and provider.** `Store/Versions.cs` holds the VersionsFile (2.4): add version, set current, rollback, lock, as-of. `Store/TableProvider.cs` implements `ITableProvider` (Latest, Locked, AsOf, Version, workspace tables, SourceOf, SaveAsVersion), with the union cache. Tests cover each truth mode and as-of across 3 versions.
- **A6. Compare** in `Ingest/Compare.cs`: `DiffOutcome Diff(Table oldT, Table newT, IReadOnlyList<string> keys)` returns the counts and the diff table. Tests: added, removed, changed, duplicate keys, key columns missing.
- **A7. Ingestion (2.5, 2.7)**:
  - `Ingest/Recipes.cs`: load, save and validate a recipe. Validation: an id slug, at least 1 folder and 1 pattern, a date column inside the signature, key columns inside the signature, and a threshold between 0 and 0.9.
  - `Ingest/Watcher.cs`.
  - `Ingest/Recognizer.cs`.
  - `Ingest/Ingestion.cs`, implementing `IIngestion`: the pipeline, quarantine, accept, reject, rollback, `undoCleanup`, the log, `seen.json`, the notice text, and the lock timer.
  - Offline tests with an in-memory `IRawSource` and a temporary `StoreRoot.ForTest` folder:
    - new month appended;
    - re-download of the previous month replaces it as v2, with v1 archived and the counts right;
    - an identical re-download is a duplicate;
    - a partial file and schema drift are quarantined, then accept and reject;
    - two periods, then accept with split;
    - future dates;
    - rollback;
    - Locked truth after the cut-off, with a post-close version stored but the locked figure unchanged;
    - as-of after 3 loads;
    - the notice sentence.
  - **The two-month pilot simulation (brief section 5) as one test:** October v1, November v1, October v2 (corrected), December v1, November v2. Zero manual steps; history and versions as expected.
- **A8. Quality (2.8)** in `Store/Quality.cs`, with a test per flag.
- **A9. History** in `Ingest/History.cs`: versions, diff (registers a `ResultStore` entry, kind `diff`) and trend.
- **A10. Workspace (2.9)**, through the queue and catalog.
- **A11. Export child and exporter (2.10)**: `ExportChild.Run` and `Store/Exporter.cs`, implementing `IExporter`. `Settings.AllowPdf` is read from `settings.json` and defaults to false.
- **A12. Commands.** `DataCommands.Register` creates the services, sets `s.Tables`, `s.Ingestion` and `s.Exporter`, loads settings, and registers every lane A command of section 4. Arguments are read with `JsonElement` helpers, and objects are serialised with `Json.Options`.
- **A13. Live tests** (`offline == false`):
  - Create `D:\ExcelAppTest\syn-ingest\in\syn-out-2026-10.xlsx` and similar files through Excel COM, from made-up values.
  - Run the real pipeline end to end through `--build`: recognise, load, re-load a corrected October, check the versions.
  - Export a result through `--export` to `D:\ExcelAppTest\syn-ingest\out\syn-pack.xlsx`, open it read-only through Excel, and check the values and the number formats.
- **Done when:**
  - Every A test passes, offline and live.
  - Group union of 12 periods × 1 M rows loads in < 3 s from warm files.
  - The pilot simulation passes.
  - No `File.Write*` outside `StoreRoot`/`ConfigRoot` exists in lane A code (grep).

### Lane B: analysis, automation and formulas (Codex GPT-6.1-Terra medium, worktree `Excel-wt-b`, branch `plan4-b`)
Lane B works against `ITableProvider` alone. Its tests use an in-memory `FakeProvider : ITableProvider` (in the test file) and made-up tables. It never needs lane A's code.

- **B1. QuerySpec and expressions.** `Query/QuerySpec.cs` mirrors `contracts.ts` exactly (Pred as a JSON object with `op`; System.Text.Json reads it by hand). `Query/Expr.cs` is a recursive-descent parser and evaluator over columns:
  - Numbers, `[Column Name]` or bare identifiers, `+ - * / ( )`, and `abs(x)`, `round(x, n)`, `coalesce(x, y)`, `pct(a, b)` = (a − b) / b × 100.
  - Nulls propagate, and division by zero gives null.
  - Relative date tokens inside Pred values: `today`, `today-N`, `startOfWeek`, `startOfLastWeek`, `startOfMonth`, `startOfLastMonth`, `endOfLastMonth`, `startOfYear`.
- **B2. Engine** in `Query/Engine.cs`: `QueryResult Run(QuerySpec spec, ITableProvider p, CancellationToken t)`. The steps are: source, joins, where (selection vector `int[]`), computed, groupBy plus aggregates (hash on composite keys: text codes and double bits), having, sort (stable, nulls last), limit, projection.
  - Text comparisons are case-insensitive.
  - An ungrouped result keeps `_v` and `_r` from the source rows, for lineage.
  - Speed, with tests using a made-up 10 M-row table: group by 2 keys with sum and count in **< 1 s**; filter plus sort of 1 M rows in < 1 s.
- **B3. Joins and fuzzy matching (idea 5)** in `Query/Join.cs` and `Query/Fuzzy.cs`. A hash join builds on the right side.
  - `normalized` keys: trim, upper case, remove spaces, `-`, `_`, `.` and `/`, and strip leading zeros from every digit run.
  - `similar`: first an exact normalised match, then, for the left keys still unmatched, the best Jaro-Winkler score over right keys of length ±2 that share their first normalised character. Accept if the score is ≥ `minScore` (default 0.92).
  - Output columns `_match` (`exact|normalized|similar|none`) and `_score`. An `anti` join lists the unmatched rows.
  - Tests: zeros, spaces, case and hyphen variants, and a 1 M × 1 M messy-key join in < 10 s.
- **B4. Results and formatting** in `Query/Format.cs`: `data.page` display text.
  - Numbers: `G15` invariant, with the decimals trimmed to 10.
  - Dates: `yyyy-MM-dd`.
  - Kinds per cell: 0 empty, 1 number, 2 text, 3 bool.
  - It works on every `ResultStore` entry, and `_` columns are never sent.
- **B5. Views, exceptions, links and drill (ideas 11, 12, 13)**:
  - Saved views live in config `views/`. An exception view counts its rows. Counts are cached and recomputed on `Ingestion.Loaded` when the view's source depends on that report type.
  - A link compiles into a left join plus computed difference and pct columns, and refreshes on `Loaded` of either side.
  - `drill` re-runs the spec without group and aggregates, with `where AND (groupCols = row values)`.
  - `lineage` uses `_v`, `_r` and `Tables.SourceOf`.
- **B6. Alerts (idea 9)**:
  - A rule is a spec (or a saved view) plus a `measure` (an aggregate name in the first row, or `rows`) plus a comparison.
  - It is evaluated right after `Loaded`, for rules whose `reportTypes` include the loaded type or are empty, and every hour.
  - `changePct` compares with the last evaluated value.
  - The same rule with the same value never fires twice in a row.
  - Fired alerts go to `alerts/log.jsonl` in the store, and emit `data.alert`.
  - A test proves an alert fires within one evaluation of a load.
- **B7. Report recipes and scheduler (ideas 8, 18)**:
  - Steps are `query`, `export` (the last result, through `s.Exporter`) and `pack`: several sheets into one workbook, one `ExportJob` with several `ExportSheet`.
  - Target tokens: `{yyyy}`, `{MM}`, `{dd}`, `{period}`, `{name}`.
  - Schedules are `daily` (time), `weekly` (day 1 = Monday, time), `monthly` (day, time) and `afterLoad` (a report type). A 30-second timer checks them. Runs missed while the app was closed run once at the next start. The last-run state is stored in config.
  - Each run emits `data.recipeRun`.
  - A test records a recipe on October data, loads November into the fake provider, and replays it: the same outputs for November.
- **B8. Sandbox (idea 15)**:
  - A sandbox is a base `TableRef` plus edits. Only edited columns are copied (copy-on-write), so creating and discarding one on 1 M rows takes < 5 s.
  - `SandboxProvider` wraps `ITableProvider` and serves `kind "sandbox"`.
  - Compare is a diff against the base, registered as `diff`.
  - Keep calls `Tables.SaveAsVersion`.
- **B9. Library (idea 10)**:
  - An `.exlib` file (JSON, `{format: "exlib1", items: [...]}`) holds views, links, alerts, report recipes, dashboards and ingest recipes (read through `s.Ingestion.Recipes`, saved through `s.Ingestion.SaveRecipe` on import).
  - Folders are replaced by `{Downloads}` and mapped back to the importer's Downloads folder.
  - `preview` lists every literal value in specs (filter values), so the user sees what leaves the machine. **No table data is ever included**; a test proves it.
  - Import renames ids on a conflict.
- **B10. Dashboards (idea 17)**:
  - Tiles compute scalars from specs. `compareTo: previousVersion` uses `TableRef.AsOf` = the last load time − 1 second. `previousPeriod` filters the recipe's date column to the previous period, through `Periods` logic in lane B's own small helper.
  - Values are cached and recomputed on `Loaded`.
- **B11. Questions (idea 20)** in `Query/Questions.cs`: rule-based, with no AI.
  - Match column names by normalised containment (best match), and table names over `Tables.List()`.
  - It must handle these 20 patterns, each tested:
    1. `top N <group> by <measure> [time]`
    2. `bottom N …`
    3. `total <measure> by <group> [time]`
    4. `count by <group>`
    5. `average <measure> by <group>`
    6. `max <measure> [by <group>]`
    7. `min …`
    8. `rows where <col> > <n>` (also `<`, `>=`, `<=`, `=`, `is`)
    9. `<col> is <value>`
    10. `show <table>`
    11. `duplicates in <col>`
    12. `blank <col>` or `missing <col>`
    13. `sort by <col> [desc|asc]`
    14. `<measure> per <group> in <period word>`
    15. `trend of <measure> by month`
    16. `compare <tableA> and <tableB> by <key>` (link)
    17. `what changed in <month name>` (history diff of the latest two versions)
    18. `how many rows in <table>`
    19. `last N days of <table>`
    20. `<measure> this month vs last month`
  - Time words: `today`, `yesterday`, `this week`, `last week`, `this month`, `last month`, `this year`, `last N days`.
  - Unknown text answers `ok: false` with a plain explanation of what was understood.
- **B12. Commands.** `QueryCommands.Register` registers every lane B command of section 4, and subscribes to `s.Ingestion.Loaded` when `s.Ingestion` is set.
- **B13. Formula engine (hybrid, brief section 1)** in `ExcelApp/Formula/`, a pure library that the lead wires into the editor:
  - `Lexer`, then `Parser`: A1 references, `$` absolute, ranges, whole columns `A:A`, sheet references `Sheet1!A1` and `'My Sheet'!A1:B9`, numbers, strings, booleans, error literals, the operators `+ - * / ^ & = <> < > <= >= %`, unary minus, and function calls.
  - The AST goes to `Evaluator`, over `interface IFormulaGrid { FValue Get(int sheet, int row, int col); int SheetIndex(string name); }`.
  - `FValue` is a number, text, bool, error or blank.
  - `DependencyGraph`:
    - Add, remove and change a formula.
    - Range dependencies are indexed per sheet and column.
    - `Recalculate(changed cells)` returns the cells to recompute in topological order.
    - A cycle marks its cells `circular`.
  - Functions:
    - Totals and counts: `SUM SUMIF SUMIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS AVERAGE AVERAGEIF AVERAGEIFS MIN MAX`.
    - Logic: `IF IFS IFERROR IFNA AND OR NOT`.
    - Lookup: `VLOOKUP HLOOKUP XLOOKUP INDEX MATCH`.
    - Maths: `ROUND ROUNDUP ROUNDDOWN ABS INT MOD`.
    - Text: `LEFT RIGHT MID LEN TRIM UPPER LOWER TEXT VALUE CONCAT CONCATENATE TEXTJOIN`.
    - Dates: `TODAY NOW DATE YEAR MONTH DAY EOMONTH EDATE WEEKDAY WEEKNUM NETWORKDAYS`.
    - `TEXT` uses the existing `NumFmt`.
    - Wildcards `* ? ~` in the `*IF(S)` and `MATCH` family.
  - An unknown function gives `FValue.Unsupported`: the editor shows the file's own value and marks the cell "Excel calculates this on save". This is the hybrid.
  - `Formula/FormulaReader.cs`: `IReadOnlyList<(int Sheet, int Row, int Col, string A1)> Read(dynamic workbook)`. Per sheet it uses `UsedRange.SpecialCells(-4123)` areas and each area's `Formula` (a 2-D array), and handles "no cells found" (COMException). It is called from inside a build child later; the lead integrates it.
  - Tests:
    - Offline: 150+ cases with expected values written from Excel's documented behaviour.
    - Live: build a made-up workbook through Excel with 60 formulas over made-up data, read Excel's values, evaluate the same formulas with the engine, and compare (numbers within 1e-9). Every mismatch is listed.
  - Speed: 100,000 formulas recalculated after one input change in < 1 s.
- **Done when:**
  - Every B test passes, offline and live.
  - The speed targets in B2, B3, B8 and B13 are met and printed.
  - The 20 question patterns all pass.
  - The library test proves no table data is exported.

### Lane C: interface (Codex GPT-6.1-Terra medium, worktree `Excel-wt-c`, branch `plan4-c`)
Lane C builds against `contracts.ts` and a complete mock (`data/mockData.ts`), so it never waits for A or B. React 19 with the existing styles. No new npm packages. Charts are SVG. Use Excel-like visual language (the existing `theme.ts`/`palette.ts`), with a left navigation rail.

- **C1. Mode switch.** In `App.tsx`, add a "Data" toggle in the title bar next to the file name; it is the only App.tsx change.
  - `mode === "data"` renders `<DataApp/>` from `data/DataApp.tsx` instead of the ribbon, grid and sheet tabs. The editor state stays mounted but hidden, so switching back loses nothing.
  - URL `?mode=data&page=<name>` opens a page directly, for screenshots.
  - `mockHost.ts`: forward any `data.*` call and event to `data/mockData.ts`.
- **C2. Shell.** `DataApp.tsx`, with:
  - a rail: Home, Data, Views, Ingestion, History, Alerts, Recipes, Sandbox, Library, Settings;
  - a top **Ask bar**;
  - a notification bell with a toast area listening to `data.notice`, `data.alert`, `data.loaded`, `data.quarantine` and `data.recipeRun`;
  - a status line with the ingestion state (`data.ingest.status`, polled every 10 s, and updated by `data.ingest` events).
- **C3. `ResultGrid.tsx`**: read-only Glide grid over `data.page` in 200-row blocks, with the same stale-reply guards as `rowStore.ts` (generation tokens).
  - Column header click: sort (rebuilds the spec).
  - Double-click a grouped row: `data.drill`, opened as a breadcrumb.
  - Alt+click a cell: `data.lineage` popover showing file, sheet, cell, version and load time ("why this number?" in two clicks).
  - `_status` rows coloured green, red or amber for added, removed and changed.
  - Right-click: Export…, Pin to Home, Save as view, Make alert.
- **C4. Data page.**
  - Projects and tables on the left (`data.workspace.*`; Add file uses the existing `pickFile` command) and report types (`data.tables.list`).
  - A table opens in ResultGrid.
  - `SpecBuilder.tsx` covers filters (Pred editor), group by, aggregates, computed columns (expression text box with validation messages from the host), joins (right table, keys, fuzzy level, min score) and sort.
  - "As of" date/version picker, and Latest/Locked toggle.
- **C5. Views page (idea 11)**: a saved views list with **exception counts as badges** (and on the rail item: the total). Open, edit, delete.
- **C6. Ingestion page (ideas 6, 7)**:
  - **Recipe wizard**, 7 steps, each with Back/Next and validation:
    1. pick a sample file (`pickFile`, then `data.ingest.preview` with progress);
    2. choose the sheet and header row (preview table);
    3. date column, granularity and week rule;
    4. key columns (with a live duplicate check on the sample);
    5. type fixes (suggested from `guesses`, toggled per column, showing before and after);
    6. safety and truth: partial threshold slider, Latest or Locked, lock working day, holidays;
    7. folders and patterns (pattern suggested from the sample name, digits replaced by `*`), then Save and "Scan now".
  - The target is that a non-technical user finishes it in under 10 minutes; keep each step to one decision.
  - **Monitor**:
    - watched folders with OK or error;
    - the queue;
    - a live stage timeline per file (noticed → recognised → compared → updated → notified);
    - the log table (`data.ingest.log`);
    - **quarantine cards** with the reason in words, plus Accept (with split and update-recipe options when relevant) and Reject;
    - each loaded record with its cleanup report (with Undo per fix), its quality card (flags with "Show rows"), and Rollback.
- **C7. History page (idea 14)**:
  - Report type, then periods, then a **version timeline** (dots with counts; locked periods marked, post-close versions marked).
  - Select two versions to get a diff in ResultGrid.
  - A **trend chart** (SVG line) of a column total across versions or periods.
  - "How did October change?" fits on one screen: timeline, change counts per version and the latest diff.
- **C8. Alerts page (idea 9)**: rule list and editor (spec or view, measure, comparison, value, report types, enabled), the fired-alerts list, and in-app toasts.
- **C9. Recipes page (ideas 8, 18, 19)**:
  - Recipe list with last run and status.
  - Step editor: query from the current spec or a saved view; export target with tokens; template file; start cell; sheet; pack of several sheets.
  - Schedule editor, Run now.
  - "Record from current view" builds a one-step recipe from the open ResultGrid's spec.
- **C10. Home (idea 17)**:
  - Dashboards as tabs, tiles in a responsive grid: value, previous value, delta % with an up or down colour, and as-of time.
  - Refreshes on `data.loaded`. Tile click opens its spec in the Data page.
  - "Pin to Home" from ResultGrid adds a tile.
  - **Opening the app in Data mode shows Home with current figures, without opening any file.**
- **C11. Sandbox page (idea 15)**:
  - Create a sandbox from the current table, and edit cells in an editable ResultGrid variant.
  - Compare shows the diff. Keep asks for report type, period and note; Discard confirms first.
- **C12. Library page (idea 10)**: pick items, then **preview the literal values** ("these values leave the machine with the file"), then export (`pickFile` save dialog; the host's existing `PickSavePath` flow; if it is not reachable, a text box for the path). Import shows the result.
- **C13. Settings**:
  - `AllowPdf`: shown disabled, with the note "waiting for security approval", unless the host says it is allowed.
  - `StartWithWindows`.
  - Read-only paths of the store and config folders.
- **C14. Ask bar (idea 20)**: type a question, then `data.ask` shows "I understood: …" with the spec as chips. Enter runs it into the Data page. `ok: false` shows the explanation and examples of the 20 patterns.
- **C15. Mock data.** `mockData.ts` answers **every** command in section 4 with made-up data, so every page works in `pnpm dev`. It covers 3 report types, 14 months, 2–3 versions per period, a quarantine, alerts, recipes and dashboards. It emits events on a timer only when `?demo=1`.
- **C16. Tests and pictures.**
  - Vitest: the spec builder produces valid QuerySpec; the Pred editor round-trips; the recipe wizard validates; the result paging guard; the toast reducer; the pattern suggestion from a file name.
  - `pnpm exec tsc --noEmit` clean. `pnpm build` passes.
  - Screenshots of every page at 1400×860 with headless Chrome (`chrome --headless=new --screenshot=<file> --window-size=1400,860 "http://localhost:5173/?mode=data&page=<page>"` against `pnpm dev`), saved in `ExcelApp/ui/.shots/data-<page>.png` (git-ignored). Inspect each picture yourself and fix what looks wrong.
- **Done when:**
  - Every page works against the mock.
  - The type check and tests are green, and the build passes.
  - 12 page pictures are inspected and correct.
  - The editor mode still works exactly as before (`vitest` old tests green; switch back and forth keeps the open file).

## 6. Merge, review and acceptance (lead and reviewer)
1. The lead commits each lane's worktree on its own branch, then merges `plan4-a`, `plan4-b` and `plan4-c` into `main`. Files are disjoint, so the merges are mechanical. Then the lead runs `check.ps1 -Offline`, `tsc`, `vitest` and `build.ps1`.
2. **The interface against the real host.** The lead runs the app, opens Data mode and walks through every page against the real engine with the made-up files from A13. Contract mismatches are fixed in the owning lane's files, by the reviewer.
3. **Reviewer: Codex GPT-6.1-Sol high, in place on `main`.** A full review of correctness, rule 0 compliance (grep for direct file writes, logs of data, COM ownership), concurrency (watcher, queue, scheduler, `Loaded` handlers), error paths and speed targets. Fix every finding test-first, then run the full live `check.ps1`, the build and the self-test.
4. **Lead: final review.**
   - Re-run everything and plant two faults: corrupt a `.part` file, and a recipe with a key not in the signature.
   - Then do the lead-owned work: editor auto-save journal and crash recovery (Phase 0), wiring the formula engine into the editor (hybrid), and the docs (README, HISTORY, DECISION-4).
   - Commit, and report to the owner.

**Acceptance (all must hold):**
- `powershell -File check.ps1` (live) all green.
- `tsc` and `vitest` green, and `build.ps1` passes.
- The pilot simulation passes.
- Speed lines printed and met:
  - group-by 10 M < 1 s;
  - a messy join of 1 M × 1 M < 10 s;
  - sandbox on 1 M rows < 5 s;
  - 100 k formulas recalculated < 1 s;
  - an alert fires within one evaluation after a load.
- A real-host walk-through of every page with pictures.
- No rule 0 violation found by grep.

## 7. Feasibility, owner decisions and reorders (brief section 7)

**Not feasible as written, and the alternative used:**
1. **DuckDB/Parquet (idea 1).** The download is blocked here (section 1). Alternative: our own columnar `.xvt` store with a C# engine. It meets the same targets (group-by over 10 M rows < 1 s with dictionary-encoded text), and DuckDB can replace it behind `ITableProvider` later.
2. **On-device AI assistant (idea 20).** There is no local model, downloads are blocked, and C: has 3.5 GB free. Alternative: the 20 rule-based patterns now, behind a small `IAssistant` seam left empty. A local model becomes possible only if IT provides one offline.
3. **"Opening any file never waits" (idea 1).** The first read of a new file still needs Excel (NASCA). Alternative: it runs in the background queue; the old data stays usable meanwhile; only brand-new files wait.
4. **PDF in the meeting pack (idea 18).** Whether NASCA protects a PDF that Excel writes is not known. Alternative: PDF is built but **off** (`AllowPdf = false`) until security confirms. The `.xlsx` packs are written by Excel and stay protected.
5. **"Visually identical" company templates (idea 19).** Needs the real approved layouts. Alternative: the template mechanism fills a copy of the approved template, which keeps its formats, merges and headers. Each template is checked against its manual report by the owner.
6. **Scheduling when the app is closed (idea 8).** Excel needs the user's session. Alternative: the engine runs while ExcelApp is open, with optional start at Windows sign-in; missed runs catch up once at start.
7. **Sharing with colleagues (idea 10).** Exporting configuration is built. Their machines need the same approval as this laptop's store before they can load data (`CachePolicy` approves one laptop). That is an IT decision.

**Owner decisions, with the defaults used until answered (brief section 6):**
1. Period and key columns per report type: chosen in the recipe wizard; no default.
2. Partial threshold: 30%.
3. How far back corrections reach: unlimited (every version kept); `KeepVersions` can limit it per recipe.
4. Latest vs Locked: Latest; lock on the 5th working day when Locked is chosen.
5. Security confirmation: the store uses the approval of 2026-11-04 for this laptop's fast copy, in the same folder and with the same lock. **The owner confirms this extension**: the store keeps more history than the fast copy.
6. Top functions and daily tasks: the formula list in B13 is the default.
7. Moving to the .NET 10 SDK, which needs about 1 GB on C: or an install on D:.

**Reorders compared with the brief:**
- The formula engine moves into this wave (lane B) instead of Phase 1, because it is a pure library with no dependency on the store.
- Ideas 13 (drill and lineage) and 4 (cleanup) become part of the foundation: lineage columns and raw copies must exist from the first load, or history cannot be traced later.
- Exports (18, 19) are built with the export child in lane A now, because recipes (8) need them.

## 8. Estimates (agent-hours of focused work; real time depends on review rounds)
| Item | Hours | Item | Hours |
|---|---|---|---|
| 1 Columnar store | 4 | 11 Exception views | 1.5 |
| 2 Workspace | 2 | 12 Live links | 1.5 |
| 3 Versioning / as-of | 3 | 13 Drill and lineage | 2 |
| 4 Typing and cleanup | 3 | 14 Restatement history | 2.5 |
| 5 Fuzzy join | 2.5 | 15 Sandbox | 2 |
| 6 Ingestion engine | 6 | 16 Quality | 2 |
| 7 Recipe wizard (UI + validation) | 4 | 17 Morning dashboard | 2.5 |
| 8 Recipes and scheduler | 3 | 18 Meeting pack | 1.5 |
| 9 Alerts | 2 | 19 Export templates | 2 |
| 10 Library | 1.5 | 20 Questions (rules) | 3 |
| Hybrid formula engine | 8 | Interface for all pages | 14 |
| Review and fixes (Sol high) | 6 | Lead integration, auto-save, docs | 5 |

## 9. The answer file (each lane, at the end, and progress after each step)
Write `D:\WORK\Software Development\GitHub\HERDR\tasks\plan4-lane-<x>-answer.md` with:
- every file created or changed;
- every check with its counts and the speed lines;
- the steps done (all, or the exact one you are on);
- deviations from this plan and why;
- contract change requests;
- problems met, with a proven reason or "not proven";
- what is not proven.

Append progress lines to `plan4-lane-<x>-progress.md` after each step.
