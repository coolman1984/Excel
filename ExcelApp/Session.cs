// The open workbook: cache, per-sheet row maps, filters, edits, find, save. Everything the interface asks for lands here.
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp;

public sealed partial class Session : IDisposable
{
    public const int MaxCells = 200_000;
    public const string Version = "1.0";
    static readonly string[] OpenExt = { ".xlsx", ".xlsm", ".xlsb", ".xls", ".csv", ".xltx", ".xltm" };

    sealed class SheetState
    {
        public FilterEngine Engine = null!; public SheetSource Source = null!;
        public int[]? Map; public int[]? Inverse; public bool Header = true; public List<(int Col, bool Desc)> Sort = new();
    }

    // ---- hooks the window (or a test) provides
    public Action<string, object?> Emit = (_, _) => { };
    public Func<string?> PickOpenPath = () => null;
    public Func<string, string?> PickSavePath = _ => null;
    public Action<string> ClipboardSet = _ => throw new UserError("There is no clipboard here.");
    public Func<string?> ClipboardGet = () => null;
    public Action<string, bool>? TitleChanged;       // file name, dirty
    public Action<string>? CloseAnswer;
    public Action? UiReady;
    public string ChildExe = System.IO.Path.Combine(AppContext.BaseDirectory, "ExcelApp.exe");
    public string DataDir;

    // ---- the open workbook
    CacheData? cache; string? path; Cells? cells; StyleSidecar? sidecar; SheetState[] sheets = Array.Empty<SheetState>();
    int current;
    public readonly EditSet Edits = new(); readonly object editLock = new();
    readonly object readerLock = new(); readonly Dictionary<int, SheetReader> readers = new();
    readonly ManualResetEventSlim gate = new(true);
    CancellationTokenSource? applyCts, findCts;
    volatile bool opening, saving; volatile bool saveOpened; Child? child;
    (int Sheet, List<(int, int)> Cells, string Hash)? pendingCut;
    FileSystemWatcher? watcher; bool changedSent; System.Threading.Timer? memTimer;
    public bool Dirty { get { lock (editLock) return Edits.Dirty; } }
    public bool Busy => opening || saving;
    public bool Saving => saving;
    public bool ChildTimedOut => child?.TimedOut ?? false;
    public string? SourcePath => path;
    public CacheData? Cache => cache;
    public string FileName => path == null ? "" : System.IO.Path.GetFileName(path);

    public Session(string? dataDir = null)
    {
        DataDir = dataDir ?? System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelApp");
        // Inserting or deleting rows or columns moves every formula's references as Excel does (EditSet calls this per formula).
        Edits.ShiftFormula = (formulaSheet, f, op) =>
        {
            var c = cache; if (c == null || op.Sheet >= c.Sheets.Count) return f;
            var d = c.Sheets[op.Sheet];
            return Formula.FormulaShift.Structural(f, formulaSheet == op.Sheet, d.Name, op.Rows, (op.Rows ? d.FirstRow : d.FirstCol) + op.At, op.Count, op.Insert);
        };
        memTimer = new System.Threading.Timer(_ => { try { double c = Machine.Commit(); if (c > 85) Emit("memory", new { CommitPercent = Math.Round(c, 1) }); } catch (Exception) { } }, null, 5000, 5000);
    }

    /// <summary>The window is closing: ask a running read to stop. Its Excel closes itself; it is never killed.</summary>
    public void CancelBuildForClose() { if (opening) child?.CloseInput(); }

    public void Dispose() { Journal?.Dispose(); Data?.Dispose(); memTimer?.Dispose(); watcher?.Dispose(); ReleaseReaders(); DisposeSheets(); }

    // =====================================================================================
    public void Register(Dictionary<string, Func<JsonElement, Task<object?>>> h)
    {
        void S(string name, Func<JsonElement, object?> f) => h[name] = a => Task.FromResult(f(a));
        S("ready", _ => { UiReady?.Invoke(); return new { version = Version, recent = Recent(), approved = CachePolicy.ThisPcApproved }; });
        S("pickFile", _ => new { path = PickOpenPath() });
        S("open", a => { Open(Arg.Str(a, "path") ?? throw new UserError("No file was chosen."), Arg.Bool(a, "discard")); return new { started = true }; });
        S("cancelOpen", _ => { child?.CloseInput(); return new { }; });
        S("workbook", _ => WorkbookInfo());
        S("sheet", a => SelectSheet(Arg.Int(a, "sheet")));
        h["rows"] = a => Task.FromResult<object?>(Rows(a));
        S("cellDetail", CellDetail);
        S("setCell", SetCell); S("clearCells", ClearCells); S("setStyle", SetStyle);
        S("undo", a => UndoRedo(true)); S("redo", a => UndoRedo(false));
        S("insertRows", a => Restructure(a, true, true)); S("deleteRows", a => Restructure(a, true, false));
        S("insertCols", a => Restructure(a, false, true)); S("deleteCols", a => Restructure(a, false, false));
        S("copy", a => Copy(a, false)); S("cut", a => Copy(a, true)); S("paste", Paste); S("fill", Fill);
        S("filterValues", FilterValues); S("setFilter", SetFilter); S("clearFilters", ClearFilters);
        S("setSort", SetSort); S("setHeaderRow", SetHeaderRow);
        S("find", Find); S("replace", Replace); S("stats", Stats);
        S("setColWidth", a => { lock (editLock) Edits.SetColWidth(Arg.Int(a, "sheet"), Arg.Int(a, "col"), Arg.Num(a, "px")); Dirtied(); return new { dirty = true }; });
        S("setFreeze", a => { lock (editLock) Edits.SetFreeze(Arg.Int(a, "sheet"), Math.Clamp(Arg.Int(a, "rows"), 0, 10), Math.Max(0, Arg.Int(a, "cols"))); Dirtied(); return new { dirty = true }; });
        h["save"] = async a => await SaveAsync(Arg.Bool(a, "saveAs"));
        S("cancelSave", _ => { if (saving && !saveOpened) child?.CloseInput(); else if (saving) throw new UserError("It is too late to cancel. Excel is already working on the file."); return new { }; });
        S("close", a => { if (Dirty && !Arg.Bool(a, "discard")) throw new UserError("Save or discard your changes before closing."); CloseWorkbook(); return new { }; });
        S("closeAnswer", a => { CloseAnswer?.Invoke(Arg.Str(a, "action") ?? "cancel"); return new { }; });
        S("recent", _ => new { recent = Recent() });
        S("recovery", _ => RecoveryInfo()); S("recoverEdits", _ => RecoverEdits()); S("discardRecovery", _ => DiscardRecovery());
        S("openExternal", _ => { if (path != null) Process.Start(new ProcessStartInfo("explorer.exe") { ArgumentList = { "/select,", path }, UseShellExecute = false, CreateNoWindow = true }); return new { }; });
        S("log", a => { string t = Arg.Str(a, "text") ?? ""; if (System.Text.RegularExpressions.Regex.IsMatch(t, @"\A(?:paint|firstScreen|jump) [0-9]{1,9}(?:\.[0-9]{1,3})? ms\z")) Log.Write("ui " + t); return new { }; });
        S("edge", Edge); S("locate", Locate); S("info", _ => Info());
        RegisterAgent(h);  // AI agents: read and write by Excel address (SessionAgent.cs)
        Agents?.Register(h); // the Agent panel: Claude Code CLI / Codex CLI
        Data?.Register(h); // plan 4: the data engine's "data.*" commands
    }

    /// <summary>Runs Claude Code CLI or Codex CLI for the Agent panel; null in tests.</summary>
    public Agent.AgentRunner? Agents;

    /// <summary>Plan 4 data engine; null in tests and the self test.</summary>
    public Data.DataHost? Data;

    // ---- recent files -------------------------------------------------------------------
    string RecentFile => System.IO.Path.Combine(DataDir, "recent.json");
    public List<string> Recent()
    {
        try { return File.Exists(RecentFile) ? JsonSerializer.Deserialize<List<string>>(File.ReadAllText(RecentFile)) ?? new() : new(); }
        catch (Exception e) when (e is IOException or JsonException) { return new(); }
    }
    void AddRecent(string p)
    {
        try
        {
            var l = Recent(); l.RemoveAll(x => string.Equals(x, p, StringComparison.OrdinalIgnoreCase)); l.Insert(0, p);
            Directory.CreateDirectory(DataDir); File.WriteAllText(RecentFile, JsonSerializer.Serialize(l.Take(12)));
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }

    // =====================================================================================
    // Opening
    // =====================================================================================
    public void Open(string file, bool discard = false)
    {
        if (Busy) throw new UserError("Please wait until the current work is finished.");
        if (Dirty && !discard) throw new UserError("Save or discard your changes before opening another file.");
        if (!File.Exists(file)) throw new UserError("That file was not found.");
        if (!OpenExt.Contains(System.IO.Path.GetExtension(file).ToLowerInvariant())) throw new UserError("This kind of file cannot be opened. Choose an Excel file.");
        if (discard && path != null && recovery == null) Journal?.Forget(path);   // the user threw these edits away
        recovery = null;
        opening = true;
        Task.Run(() => OpenCore(System.IO.Path.GetFullPath(file)));
    }

    public Task OpenAsyncForTest(string file) { Open(file); return Task.Run(async () => { while (opening) await Task.Delay(20); }); }

    void OpenCore(string file)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            if (!CachePolicy.ThisPcApproved) { OpenFailed("The fast copy is approved for one laptop only, and this is not that laptop."); return; }
            string cp = CachePolicy.PathFor(file);
            Emit("opening", new { Phase = "Checking the file", Percent = 0, Seconds = 0.0 });
            if (!CacheData.HeaderMatches(cp, file))
            {
                double commit = Machine.Commit();
                if (commit >= Machine.StopAt) { OpenFailed($"Windows memory is at {commit:F0}%. Nothing was started. Close some programs and try again."); return; }
                if (cache != null && string.Equals(path, file, StringComparison.OrdinalIgnoreCase)) { ReleaseReaders(); }
                int code = RunBuild(file, sw);
                if (code != 0)
                {
                    OpenFailed(code switch
                    {
                        2 => "Cancelled. Excel was closed and no copy was kept.",
                        90 => $"Stopped because Windows memory reached {Machine.StopAt:F0}%. Close some programs and try again.",
                        5 => "There is not enough free disk space for the copy. Nothing was kept.",
                        _ => "Excel could not open this file. If it needs a password, that is not supported.",
                    });
                    return;
                }
            }
            LoadCache(file, cp, -1);
            AddRecent(file);
            Log.Write($"opened sheets={cache!.Sheets.Count} cells={cache.Sheets.Sum(s => s.Cells)} {sw.ElapsedMilliseconds} ms");
            opening = false;
            Emit("opened", new { Workbook = WorkbookInfo() });
            FindRecovery(file);
        }
        catch (InvalidDataException) { OpenFailed("The fast copy could not be read."); }
        catch (CachePolicyException e) { OpenFailed(e.Message); }
        catch (Exception e) { Log.Write("open failed " + e.GetType().Name); OpenFailed($"The file could not be opened ({e.GetType().Name})."); }
        finally { opening = false; }
    }

    void OpenFailed(string msg) { Emit("openFailed", new { Error = msg }); }

    sealed class Child
    {
        public Process P = null!; public TaskCompletionSource<int> Done = new();
        public volatile bool TimedOut;
        public void CloseInput() { try { P.StandardInput.Close(); } catch (Exception e) when (e is IOException or InvalidOperationException) { } }
    }

    Child StartChild(string mode, string? arg, Action<string> onLine)
    {
        var psi = new ProcessStartInfo(ChildExe) { RedirectStandardOutput = true, RedirectStandardInput = true, UseShellExecute = false, CreateNoWindow = true };
        psi.ArgumentList.Add(mode); if (arg != null) psi.ArgumentList.Add(arg);
        var c = new Child { P = new Process { StartInfo = psi, EnableRaisingEvents = true } };
        c.P.OutputDataReceived += (s, e) => { if (e.Data != null) try { onLine(e.Data); } catch (Exception ex) { Log.Write("line handler " + ex.GetType().Name); } };
        c.P.Exited += (s, e) => { try { c.P.WaitForExit(); } catch (Exception) { } c.Done.TrySetResult(c.P.ExitCode); };
        c.P.Start(); c.P.BeginOutputReadLine();
        return c;
    }

    async Task<int> WaitChild(Child running, bool save)
    {
        if (await Task.WhenAny(running.Done.Task, Task.Delay(TimeSpan.FromMinutes(5))) != running.Done.Task) {
            running.TimedOut = true;
            const string message = "Excel is taking longer than expected. You can close this window. Its own Excel will be left to finish safely.";
            Emit(save ? "saving" : "opening", new { Phase = message, Percent = 0, CanCancel = false });
            Log.Write("child exceeded time limit");
        }
        return await running.Done.Task;
    }

    int RunBuild(string file, Stopwatch sw)
    {
        var c = StartChild("--build", file, line =>
        {
            double secs = sw.Elapsed.TotalSeconds;
            if (line == "OPENING") Emit("opening", new { Phase = "Starting Excel", Percent = 0, CanCancel = true, Seconds = secs });
            else if (line is "OPENED" or "FORMATS") Emit("opening", new { Phase = "Reading number formats", Percent = 0, Seconds = secs });
            else if (line.StartsWith("READ ") && int.TryParse(line.AsSpan(5), out int p)) Emit("opening", new { Phase = "Making the fast copy", Percent = p, Seconds = secs });
            else if (line == "CHECKING") Emit("opening", new { Phase = "Checking numbers against Excel", Percent = 100, Seconds = secs });
        });
        child = c;
        try { return WaitChild(c, false).GetAwaiter().GetResult(); } finally { child = null; }
    }

    void LoadCache(string file, string cachePath, int keepSheet)
    {
        findCts?.Cancel();
        var data = CacheData.Load(cachePath);
        ReleaseReaders(); DisposeSheets();
        lock (editLock) Edits.Clear();
        cache = data; path = file; pendingCut = null;
        sidecar = null; try { sidecar = StyleStore.Load(file, data); } catch (Exception e) { Log.Write("sidecar " + e.GetType().Name); }
        cells = new Cells(data, Edits, editLock, () => sidecar);
        sheets = new SheetState[data.Sheets.Count];
        for (int i = 0; i < sheets.Length; i++) sheets[i] = NewState(i);
        current = keepSheet >= 0 && keepSheet < sheets.Length ? keepSheet : Math.Max(0, data.Sheets.FindIndex(s => !s.Hidden));
        Watch(file); changedSent = false;
        TitleChanged?.Invoke(FileName, false);
    }

    SheetState NewState(int i)
    {
        var st = new SheetState { Source = new SheetSource(cache!, i, cells!) };
        st.Engine = new FilterEngine(st.Source); return st;
    }

    void DisposeSheets() { foreach (var s in sheets) s.Source.Dispose(); sheets = Array.Empty<SheetState>(); }

    void Watch(string file)
    {
        watcher?.Dispose(); watcher = null;
        try
        {
            watcher = new FileSystemWatcher(System.IO.Path.GetDirectoryName(file)!, System.IO.Path.GetFileName(file)) { NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size, EnableRaisingEvents = true };
            watcher.Changed += (s, e) =>
            {
                if (saving || opening || changedSent || cache == null) return;
                Thread.Sleep(300);
                try { if (!cache.IsFreshFor(file)) { changedSent = true; Emit("fileChanged", null); } } catch (IOException) { }
            };
        }
        catch (Exception e) when (e is ArgumentException or IOException) { watcher = null; }
    }

    void CloseWorkbook()
    {
        if (Busy) throw new UserError("Please wait until the current work is finished.");
        findCts?.Cancel();
        ReleaseReaders(); DisposeSheets(); watcher?.Dispose(); watcher = null;
        if (path != null && recovery == null) Journal?.Forget(path);   // closing means saved or discarded; a waiting recovery stays for next time
        recovery = null;
        cache = null; path = null; cells = null; sidecar = null; lock (editLock) Edits.Clear();
        TitleChanged?.Invoke("", false);
    }

    // ---- readers --------------------------------------------------------------------------
    SheetReader Reader(int sheet)
    {
        if (!readers.TryGetValue(sheet, out var r)) readers[sheet] = r = new SheetReader(cache!, cache!.Sheets[sheet]);
        return r;
    }
    void ReleaseReaders() { lock (readerLock) { foreach (var r in readers.Values) r.Dispose(); readers.Clear(); } foreach (var s in sheets) s.Source.Release(); cells?.Calc.Release(); }

    CacheData Need() => cache ?? throw new UserError("No file is open.");
    SheetState St(int sheet) { Need(); if (sheet < 0 || sheet >= sheets.Length) throw new UserError("That sheet does not exist."); return sheets[sheet]; }
    int ViewRows(CacheData c, SheetState s, int sheet) => s.Map?.Length ?? RowsOf(sheet);
    /// <summary>The sheet's size as the editor shows it: the fast copy's, plus inserted and minus deleted rows and columns.</summary>
    int RowsOf(int sheet) { lock (editLock) return Edits.RowCount(sheet, cache!.Sheets[sheet].Rows); }
    int ColsOf(int sheet) { lock (editLock) return Edits.ColCount(sheet, cache!.Sheets[sheet].Cols); }
    static int Real(SheetState s, int view) => s.Map != null ? s.Map[view] : view;
    int ToView(int sheet, int real)
    {
        var s = sheets[sheet];
        if (real < 0 || real >= RowsOf(sheet)) return -1;
        if (s.Map == null) return real;
        if (s.Inverse == null) { var inv = new int[RowsOf(sheet)]; Array.Fill(inv, -1); for (int i = 0; i < s.Map.Length; i++) inv[s.Map[i]] = i; s.Inverse = inv; }
        return real >= 0 && real < s.Inverse.Length ? s.Inverse[real] : -1;
    }

    // =====================================================================================
    // Workbook and sheets
    // =====================================================================================
    public object WorkbookInfo()
    {
        var c = Need();
        var list = c.Sheets.Select((d, i) =>
        {
            int rows = RowsOf(i), cols = ColsOf(i);
            var widths = new double[cols];
            for (int k = 0; k < cols; k++)
            {
                int pc; lock (editLock) pc = Edits.PhysCol(i, k);
                widths[k] = pc >= 0 && pc < d.ColWidths.Length ? Math.Round(d.ColWidths[pc] * 7 + 5) : 64;   // an inserted column: Excel's default width
                lock (editLock) if (Edits.ColWidths.TryGetValue((i, k), out var px)) widths[k] = px;
            }
            int fr = 0, fc = 0; if (sidecar != null && sidecar.Sheets.TryGetValue(i, out var s2)) { fr = s2.FrozenRows; fc = s2.FrozenCols; }
            lock (editLock) if (Edits.Freeze.TryGetValue(i, out var f)) { fr = f.Rows; fc = f.Cols; }
            return new { index = i, name = d.Name, hidden = d.Hidden, rows, cols, firstRow = d.FirstRow, firstCol = d.FirstCol, colWidthsPx = widths, frozenRows = fr, frozenCols = fc };
        }).ToList();
        return new { name = FileName, path, sheets = list, dirty = Dirty, currentSheet = current, partialStyles = sidecar?.Partial ?? false };
    }

    object SelectSheet(int sheet)
    {
        var s = St(sheet); current = sheet; return SheetSummary(sheet, s);
    }

    object SheetSummary(int sheet, SheetState s) => new
    {
        viewRows = ViewRows(cache!, s, sheet), filterActive = s.Engine.Conditions.Count > 0, header = s.Header,
        sort = s.Sort.Select(k => new { col = k.Col, desc = k.Desc }).ToList(),
        filters = s.Engine.Conditions.ToDictionary(kv => kv.Key.ToString(CultureInfo.InvariantCulture), kv => ConditionWire(kv.Value)),
    };

    static object ConditionWire(Condition c) => c switch
    {
        ValuesCondition v => new { kind = "values", include = v.Include.ToList(), blanks = v.Blanks },
        NumberCondition n => new { kind = "number", op = n.Op, a = n.A, b = n.B },
        TextCondition t => new { kind = "text", op = t.Op, value = t.Value },
        _ => new { kind = "none" },
    };

    object Info()
    {
        var c = Need(); long size = 0; bool nasca = false;
        try
        {
            var f = new FileInfo(path!); size = f.Length;
            using var fs = f.OpenRead(); var b = new byte[8]; int n = fs.Read(b, 0, 8); nasca = n == 8 && Encoding.ASCII.GetString(b) == "<## NASC";
        }
        catch (IOException) { }
        long cacheBytes = 0; try { cacheBytes = new FileInfo(c.Path).Length; } catch (IOException) { }
        return new { path, size, cacheBytes, nasca, readTime = new DateTime(c.SourceTicks, DateTimeKind.Utc).ToLocalTime().ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) };
    }

    // =====================================================================================
    // Reading cells
    // =====================================================================================
    object Rows(JsonElement a)
    {
        gate.Wait();
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        int total = ViewRows(c, st, sheet), from = Math.Max(0, Arg.Int(a, "from")), count = Math.Clamp(Arg.Int(a, "count"), 0, 500);
        int colFrom = Math.Clamp(Arg.Int(a, "colFrom", 0), 0, Math.Max(0, ColsOf(sheet))), colCount = Math.Clamp(Arg.Int(a, "colCount", ColsOf(sheet)), 0, ColsOf(sheet) - colFrom);
        count = Math.Min(count, Math.Max(1, 400_000 / Math.Max(1, colCount)));
        count = Math.Max(0, Math.Min(count, total - from));
        var map = st.Map; var rows = new object[count];
        lock (readerLock)
        {
            var rd = Reader(sheet);
            for (int i = 0; i < count; i++)
            {
                int real = map != null ? map[from + i] : from + i;
                var cs = new object?[colCount];
                for (int k = 0; k < colCount; k++) cs[k] = cells!.Wire(cells.Render(sheet, rd, real, colFrom + k));
                rows[i] = new object[] { d.FirstRow + real, cs };
            }
        }
        return new { from, rows };
    }

    object CellDetail(JsonElement a)
    {
        gate.Wait();
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        int vr = Arg.Int(a, "row"), col = Arg.Int(a, "col"); CheckCell(c, st, sheet, vr, col);
        int real = Real(st, vr);
        lock (readerLock)
        {
            var rd = Reader(sheet); var s = cells!.Render(sheet, rd, real, col);
            bool edited = cells.TryEdit(sheet, real, col, out _);
            return new
            {
                address = Cells.Address(d.FirstRow, d.FirstCol, real, col), raw = cells.Raw(sheet, rd, real, col), edited,
                style = Cells.StyleJson(s.Style), format = cells.CurrentFormat(sheet, real, col), kind = (int)s.Kind,
            };
        }
    }

    void CheckCell(CacheData c, SheetState st, int sheet, int vr, int col)
    {
        if (vr < 0 || vr >= ViewRows(c, st, sheet) || col < 0 || col >= ColsOf(sheet)) throw new UserError("That cell is outside the sheet.");
    }

    // =====================================================================================
    // Editing
    // =====================================================================================
    void Dirtied() { TitleChanged?.Invoke(FileName, Dirty); ScheduleJournal(); }

    // ---- auto-save journal (EditJournal) -------------------------------------------------
    /// <summary>Set by the window when the store is available; null means no journal (tests, or not the approved laptop).</summary>
    public EditJournal? Journal;
    EditJournal.Snapshot? recovery;   // unsaved edits found when the file was opened, waiting for the user's answer

    void ScheduleJournal()
    {
        var j = Journal; string? p = path;
        if (j == null || p == null || recovery != null) return;   // while a recovery waits, the old journal must not be overwritten
        j.Schedule(() => { lock (editLock) return string.Equals(path, p, StringComparison.OrdinalIgnoreCase) ? EditJournal.Take(p, Edits) : null; });
    }

    void FindRecovery(string file)
    {
        recovery = Journal?.Find(file);
        if (recovery != null) Emit("recovery", RecoveryInfo());
    }

    object RecoveryInfo() => recovery == null ? new { available = false } :
        new { available = true, cells = recovery.Cells.Count, savedAt = recovery.SavedAt.ToString("yyyy-MM-dd HH:mm"), sameFile = EditJournal.SameFile(recovery) };

    object RecoverEdits()
    {
        var c = Need(); var r = recovery ?? throw new UserError("There are no unsaved changes to restore.");
        int n; lock (editLock) n = EditJournal.Restore(r, Edits, s => s >= 0 && s < c.Sheets.Count ? (c.Sheets[s].Rows, c.Sheets[s].Cols) : (-1, -1));
        recovery = null;
        Dirtied();
        return new { restored = n, dirty = Dirty, canUndo = Edits.CanUndo };
    }

    object DiscardRecovery()
    {
        if (recovery != null) { Journal?.Forget(recovery.File); recovery = null; }
        return new { };
    }

    // View-space rectangles clamped to the sheet, as a flat list.
    List<(int R, int C)> Expand(CacheData c, SheetState st, int sheet, List<(int R1, int C1, int R2, int C2)> ranges)
    {
        int vrows = ViewRows(c, st, sheet), cols = ColsOf(sheet); long n = 0;
        var rects = ranges.Select(r => (R1: Math.Max(0, r.R1), C1: Math.Max(0, r.C1), R2: Math.Min(vrows - 1, r.R2), C2: Math.Min(cols - 1, r.C2))).Where(r => r.R1 <= r.R2 && r.C1 <= r.C2).ToList();
        foreach (var r in rects) n += (long)(r.R2 - r.R1 + 1) * (r.C2 - r.C1 + 1);
        if (n > MaxCells) throw new UserError($"Too many cells are selected. The limit is {MaxCells:N0} cells.");
        var list = new List<(int, int)>((int)n);
        foreach (var r in rects) for (int i = r.R1; i <= r.R2; i++) for (int k = r.C1; k <= r.C2; k++) list.Add((i, k));
        return list;
    }

    object Mutate(int sheet, SheetState st, List<(int Row, int Col, CellEdit? After)> realEdits)
    {
        List<(int Row, int Col)> changed;
        lock (editLock) changed = Edits.Apply(sheet, realEdits);
        Dirtied();
        return MutationResult(sheet, changed);
    }

    object MutationResult(int sheet, List<(int Row, int Col)> changed)
    {
        var list = new List<int[]>(changed.Count);
        foreach (var (r, k) in changed) { int v = ToView(sheet, r); if (v >= 0) list.Add(new[] { v, k }); }
        // recalc: formulas may show new results anywhere, so the interface reads the visible rows again.
        return new { sheet, changed = list, dirty = Dirty, canUndo = Edits.CanUndo, canRedo = Edits.CanRedo, recalc = cells?.HasFormulas ?? false };
    }

    CellEdit KeepStyle(int sheet, int real, int col, CellEdit fresh)
    {
        if (!cells!.TryEdit(sheet, real, col, out var old) || old.Style == null) return fresh;
        return fresh with { Style = fresh.Style == null ? old.Style : old.Style.Merge(fresh.Style) };
    }

    CellEdit Typed(int sheet, int real, int col, string input)
    {
        var parsed = InputParser.Parse(input, CultureInfo.CurrentCulture, cells!.CurrentFormat(sheet, real, col), cache!.Date1904, out _);
        if (parsed.Kind == EditKind.Formula)
        {
            // As Excel does: close brackets left open, and refuse a formula that cannot be read.
            string f = Formula.Calc.Complete(parsed.Formula!.Trim());
            if (Formula.Calc.Problem(f) is string problem) throw new UserError(problem);
            parsed = parsed with { Formula = f };
        }
        return KeepStyle(sheet, real, col, parsed);
    }

    object SetCell(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet);
        int vr = Arg.Int(a, "row"), col = Arg.Int(a, "col"); CheckCell(c, st, sheet, vr, col);
        int real = Real(st, vr);
        var edit = Typed(sheet, real, col, Arg.Str(a, "input") ?? "");
        return Mutate(sheet, st, new() { (real, col, edit) });
    }

    object ClearCells(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet);
        string what = Arg.Str(a, "what") ?? "contents";
        var list = new List<(int, int, CellEdit?)>();
        foreach (var (vr, col) in Expand(c, st, sheet, Arg.Ranges(a)))
        {
            int real = Real(st, vr); cells!.TryEdit(sheet, real, col, out var old);
            CellEdit? after;
            if (what == "formats") after = old == null ? null : EditMarks.IsStyleOnly(old) ? null : old with { Style = null };
            else if (what == "all") after = new CellEdit(EditKind.Clear);
            else after = old?.Style != null ? new CellEdit(EditKind.Clear, Style: old.Style) : new CellEdit(EditKind.Clear);
            list.Add((real, col, after));
        }
        return Mutate(sheet, st, list);
    }

    object SetStyle(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet);
        var delta = ParseStyle(a.GetProperty("style"));
        var list = new List<(int, int, CellEdit?)>();
        foreach (var (vr, col) in Expand(c, st, sheet, Arg.Ranges(a)))
        {
            int real = Real(st, vr);
            CellEdit after = cells!.TryEdit(sheet, real, col, out var old)
                ? old with { Style = (old.Style ?? new StyleDelta()).Merge(delta) }
                : new CellEdit(EditKind.Clear, Text: EditMarks.StyleOnly, Style: delta);
            list.Add((real, col, after));
        }
        return Mutate(sheet, st, list);
    }

    static int ParseColor(JsonElement v)
    {
        if (v.ValueKind == JsonValueKind.Number) return v.GetInt32() & 0xFFFFFF;
        string s = (v.GetString() ?? "").TrimStart('#');
        return int.Parse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture) & 0xFFFFFF;
    }

    public static StyleDelta ParseStyle(JsonElement s)
    {
        bool? B(string n) => s.TryGetProperty(n, out var v) && v.ValueKind is JsonValueKind.True or JsonValueKind.False ? v.GetBoolean() : null;
        int? Col(string n) => s.TryGetProperty(n, out var v) && v.ValueKind is JsonValueKind.String or JsonValueKind.Number ? ParseColor(v) : null;
        int? align = s.TryGetProperty("align", out var al) && al.ValueKind == JsonValueKind.Number ? al.GetInt32() : null;
        string? border = Arg.Str(s, "border");
        if (border is not (null or "none" or "all" or "outside" or "bottom")) throw new UserError("That border is not known.");
        return new StyleDelta(B("bold"), B("italic"), B("underline"), Col("fontColor"), Col("fill"), align, Arg.Str(s, "numberFormat"), border);
    }

    object UndoRedo(bool undo)
    {
        Need(); (int Sheet, List<(int, int)> Changed) r;
        bool structural;
        lock (editLock) { r = undo ? Edits.Undo() : Edits.Redo(); structural = r.Sheet >= 0 && Edits.LastWasStructural; }
        Dirtied();
        if (r.Sheet < 0) return new { sheet = -1, changed = new List<int[]>(), dirty = Dirty, canUndo = Edits.CanUndo, canRedo = Edits.CanRedo };
        if (structural) { AfterStructure(r.Sheet); return StructureResult(r.Sheet); }
        return MutationResult(r.Sheet, r.Changed);
    }

    // ---- inserting and deleting rows and columns ----------------------------------------------
    /// <summary>insertRows / deleteRows {sheet, row, count} and insertCols / deleteCols {sheet, col, count}, in the
    /// interface's coordinates. One undo step; saved by replaying it in Excel (SaveBack).</summary>
    object Restructure(JsonElement a, bool rows, bool insert)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        if (st.Map != null) throw new UserError("Clear the filter and the sorting on this sheet first, then insert or delete.");
        int at = Arg.Int(a, rows ? "row" : "col"), count = Arg.Int(a, "count", 1);
        int size = rows ? RowsOf(sheet) : ColsOf(sheet), max = rows ? 1_048_576 - d.FirstRow + 1 : 16_384 - d.FirstCol + 1;
        string what = rows ? (count == 1 ? "row" : "rows") : (count == 1 ? "column" : "columns");
        if (count < 1 || count > max) throw new UserError($"Choose between 1 and {max:N0} {what}.");
        if (insert)
        {
            if (at < 0 || at > size) throw new UserError("That place is outside the sheet.");
            if (size + count > max) throw new UserError($"There is no room for {count:N0} more {what}: Excel sheets stop at {(rows ? "row 1,048,576" : "column XFD")}.");
        }
        else
        {
            if (at < 0 || at + count > size) throw new UserError($"Those {what} are outside the sheet.");
            if (count >= size) throw new UserError($"A sheet keeps at least one {(rows ? "row" : "column")}.");
        }
        lock (editLock) Edits.Structure(new StructOp(sheet, rows, at, count, insert), d.Rows, d.Cols);
        AfterStructure(sheet);
        Dirtied();
        return StructureResult(sheet);
    }

    /// <summary>After an insert or delete (or its undo): the view is the plain sheet again, and cut or copied positions are gone.</summary>
    void AfterStructure(int sheet)
    {
        var st = sheets[sheet];
        st.Engine.ClearConditions(); st.Sort = new(); st.Engine.SetSort(st.Sort); st.Map = null; st.Inverse = null;
        pendingCut = null; copiedFormulas = null;
    }

    object StructureResult(int sheet) => new
    {
        sheet, structure = true, rows = RowsOf(sheet), cols = ColsOf(sheet), viewRows = ViewRows(cache!, sheets[sheet], sheet),
        dirty = Dirty, canUndo = Edits.CanUndo, canRedo = Edits.CanRedo, recalc = cells?.HasFormulas ?? false,
    };

    // ---- clipboard ------------------------------------------------------------------------
    static string Hash(string s) => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(s)));

    object Copy(JsonElement a, bool cut)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet);
        var ranges = Arg.Ranges(a); long n = 0;
        foreach (var r in ranges) n += (long)(r.R2 - r.R1 + 1) * (r.C2 - r.C1 + 1);
        if (n > CopyText.MaxCells) throw new UserError($"Too many cells are selected. Copy works up to {CopyText.MaxCells:N0} cells.");
        var list = Expand(c, st, sheet, ranges);
        if (list.Count == 0) throw new UserError("Nothing is selected.");
        string text;
        lock (readerLock)
        {
            var rd = Reader(sheet);
            text = CopyText.Build(list, (vr, col) => cells!.Render(sheet, rd, Real(st, vr), col).Text);
        }
        ClipboardSet(text);
        pendingCut = cut ? (sheet, list.Select(x => (Real(st, x.R), x.C)).ToList(), Hash(text)) : null;
        // Remember the formulas so a paste of this same text pastes formulas, as Excel does (one rectangle only).
        copiedFormulas = null;
        if (ranges.Count == 1)
        {
            var r0 = ranges[0]; var map = new Dictionary<(int, int), (string F, int Row, int Col)>();
            foreach (var (vr, col) in list)
                if (cells!.TryEdit(sheet, Real(st, vr), col, out var e) && e.Kind == EditKind.Formula)
                    map[(vr - Math.Max(0, r0.R1), col - Math.Max(0, r0.C1))] = (e.Formula ?? "", Real(st, vr), col);
            if (map.Count > 0) copiedFormulas = (Hash(text), map);
        }
        return new { cells = list.Count };
    }

    (string Hash, Dictionary<(int, int), (string F, int Row, int Col)> Map)? copiedFormulas;

    /// <summary>Splits clipboard text on tabs and line breaks; a quoted cell may hold both. The inverse of CopyText's quoting.</summary>
    public static List<List<string>> SplitClipboard(string text)
    {
        var rows = new List<List<string>>(); var row = new List<string>(); var sb = new StringBuilder(); int i = 0, cells = 0; bool any = false;
        void AddCell() {
            if (++cells > MaxCells) throw new UserError($"Too much to paste. The limit is {MaxCells:N0} cells.");
            row.Add(sb.ToString()); sb.Clear();
        }
        while (i < text.Length)
        {
            char ch = text[i];
            if (ch == '"' && sb.Length == 0 && (i == 0 || text[i - 1] is '\t' or '\n'))
            {
                i++;
                while (i < text.Length) { if (text[i] == '"') { if (i + 1 < text.Length && text[i + 1] == '"') { sb.Append('"'); i += 2; continue; } i++; break; } sb.Append(text[i++]); }
                any = true; continue;
            }
            if (ch == '\t') { AddCell(); any = true; i++; }
            else if (ch == '\r' || ch == '\n') { AddCell(); rows.Add(row); row = new(); any = false; i += ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n' ? 2 : 1; }
            else { sb.Append(ch); any = true; i++; }
        }
        if (any || sb.Length > 0 || row.Count > 0) { AddCell(); rows.Add(row); }
        return rows;
    }

    object Paste(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        int row0 = Arg.Int(a, "row"), col0 = Arg.Int(a, "col");
        CheckCell(c, st, sheet, row0, col0);
        string text = ClipboardGet() ?? "";
        if (text.Length == 0) throw new UserError("The clipboard has no text to paste.");
        var grid = SplitClipboard(text); long n = grid.Sum(r => (long)r.Count);
        if (n > MaxCells) throw new UserError($"Too much to paste. The limit is {MaxCells:N0} cells.");
        int vrows = ViewRows(c, st, sheet); var list = new List<(int, int, CellEdit?)>(); var seen = new HashSet<(int, int)>();
        string hash = Hash(text);
        var formulas = copiedFormulas is { } cf && cf.Hash == hash ? cf.Map : null;
        bool moving = pendingCut is { } pc && pc.Sheet == sheet && pc.Hash == hash;
        for (int i = 0; i < grid.Count; i++)
            for (int k = 0; k < grid[i].Count; k++)
            {
                int vr = row0 + i, col = col0 + k; if (vr >= vrows || col >= ColsOf(sheet)) continue;
                int real = Real(st, vr); seen.Add((real, col));
                if (formulas != null && formulas.TryGetValue((i, k), out var src))
                {
                    string f = moving ? src.F : Formula.FormulaShift.Shift(src.F, real - src.Row, col - src.Col);
                    list.Add((real, col, KeepStyle(sheet, real, col, new CellEdit(EditKind.Formula, Formula: f))));
                    continue;
                }
                list.Add((real, col, Typed(sheet, real, col, grid[i][k])));
            }
        if (pendingCut is { } otherCut && otherCut.Sheet != sheet && otherCut.Hash == Hash(text))
            throw new UserError("Cut cells can only be moved within the same sheet. Use Copy to paste on another sheet.");
        if (pendingCut is { } cut && cut.Sheet == sheet && cut.Hash == Hash(text))
            foreach (var (r, k) in cut.Cells) if (!seen.Contains((r, k))) list.Add((r, k, cells!.TryEdit(sheet, r, k, out var old) && old.Style != null ? new CellEdit(EditKind.Clear, Style: old.Style) : new CellEdit(EditKind.Clear)));
        pendingCut = null;
        return Mutate(sheet, st, list);
    }

    object Fill(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        var src = Arg.Rect(a.GetProperty("from")); var dst = Arg.Rect(a.GetProperty("to"));
        int h = src.R2 - src.R1 + 1, w = src.C2 - src.C1 + 1; int vrows = ViewRows(c, st, sheet);
        if (src.R1 < 0 || src.C1 < 0 || src.R2 >= vrows || src.C2 >= ColsOf(sheet)) throw new UserError("That fill is outside the sheet.");
        long n = (long)(Math.Min(dst.R2, vrows - 1) - Math.Max(dst.R1, 0) + 1) * (Math.Min(dst.C2, ColsOf(sheet) - 1) - Math.Max(dst.C1, 0) + 1);
        if (n > MaxCells) throw new UserError($"Too many cells are selected. The limit is {MaxCells:N0} cells.");
        var list = new List<(int, int, CellEdit?)>();
        // Filling down or up runs along rows; filling left or right runs along columns. Each line of the source makes its own series.
        bool vertical = dst.C1 >= src.C1 && dst.C2 <= src.C2;
        bool copyOnly = Arg.Bool(a, "copy");   // Ctrl+D and Ctrl+R copy exactly; the fill handle continues a series
        var series = new Dictionary<int, FillSeries>();
        lock (readerLock)
        {
            var rd = Reader(sheet);
            FillSeries Line(int line)
            {
                if (series.TryGetValue(line, out var fs)) return fs;
                int n = vertical ? h : w; var items = new CellEdit?[n]; var formats = new string[n];
                for (int i = 0; i < n; i++)
                {
                    int r = Real(st, vertical ? src.R1 + i : line), c = vertical ? line : src.C1 + i;
                    items[i] = cells!.ValueEdit(sheet, rd, r, c); formats[i] = cells.CurrentFormat(sheet, r, c);
                }
                return series[line] = FillSeries.Of(items, formats);
            }
            for (int vr = Math.Max(dst.R1, 0); vr <= Math.Min(dst.R2, vrows - 1); vr++)
                for (int col = Math.Max(dst.C1, 0); col <= Math.Min(dst.C2, ColsOf(sheet) - 1); col++)
                {
                    if (vr >= src.R1 && vr <= src.R2 && col >= src.C1 && col <= src.C2) continue;
                    int sr = src.R1 + (((vr - src.R1) % h) + h) % h, sc = src.C1 + (((col - src.C1) % w) + w) % w;
                    int realS = Real(st, sr), real = Real(st, vr);
                    var v = cells!.ValueEdit(sheet, rd, realS, sc) ?? new CellEdit(EditKind.Clear);
                    if (v.Kind == EditKind.Formula) v = v with { Formula = Formula.FormulaShift.Shift(v.Formula ?? "", real - realS, col - sc) };
                    else if (!copyOnly)
                    {
                        int t = vertical ? vr - src.R1 : col - src.C1;
                        v = Line(vertical ? col : vr).At(t) ?? v;
                    }
                    cells.TryEdit(sheet, realS, sc, out var srcEdit);
                    v = v with { Style = srcEdit?.Style };
                    list.Add((real, col, KeepStyle(sheet, real, col, v)));
                }
        }
        return Mutate(sheet, st, list);
    }

    // =====================================================================================
    // Filter and sort
    // =====================================================================================
    static Condition? ParseCondition(JsonElement? e)
    {
        if (e == null) return null; var j = e.Value;
        switch (Arg.Str(j, "kind"))
        {
            case "values":
                var set = new HashSet<string>(StringComparer.Ordinal);
                if (j.TryGetProperty("include", out var inc) && inc.ValueKind == JsonValueKind.Array) foreach (var t in inc.EnumerateArray()) set.Add(t.GetString() ?? "");
                return new ValuesCondition(set, Arg.Bool(j, "blanks"));
            case "number":
                return new NumberCondition(Arg.Str(j, "op") ?? "eq", Arg.Num(j, "a", 0), j.TryGetProperty("b", out var b) && b.ValueKind == JsonValueKind.Number ? b.GetDouble() : null);
            case "text":
                return new TextCondition(Arg.Str(j, "op") ?? "eq", Arg.Str(j, "value") ?? "");
            default: throw new UserError("That filter is not known.");
        }
    }

    object FilterValues(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); int col = Arg.Int(a, "col");
        if (col < 0 || col >= ColsOf(sheet)) throw new UserError("That column does not exist.");
        var (values, blanks, truncated) = st.Engine.Distinct(col, Arg.Str(a, "search"));
        // Number filters for columns that are mostly numbers, text filters otherwise.
        int rows = RowsOf(sheet), from = Math.Min(rows, st.Header ? 1 : 0), n = Math.Min(1000, rows - from), nums = 0, texts = 0;
        if (n > 0) { var t = new string[n]; var k = new byte[n]; var m = new double[n]; st.Source.Read(col, from, n, t, k, m); foreach (var x in k) { if (x == 1) nums++; else if (x == 2) texts++; } }
        return new { values = values.Select(v => new { text = v.Text, count = v.Count }).ToList(), blanks, truncated, kind = nums >= texts ? "number" : "text" };
    }

    object RecomputeView(int sheet, SheetState st)
    {
        applyCts?.Cancel(); var cts = applyCts = new CancellationTokenSource();
        var sw = Stopwatch.StartNew();
        st.Engine.HeaderRows = st.Header ? 1 : 0;
        var map = st.Engine.Apply(cts.Token);
        st.Map = map; st.Inverse = null;
        Log.Write($"view rebuilt rows={ViewRows(cache!, st, sheet)} {sw.ElapsedMilliseconds} ms");
        return new { viewRows = ViewRows(cache!, st, sheet), filterActive = st.Engine.Conditions.Count > 0 };
    }

    object SetFilter(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); int col = Arg.Int(a, "col");
        if (col < 0 || col >= ColsOf(sheet)) throw new UserError("That column does not exist.");
        st.Engine.SetCondition(col, ParseCondition(Arg.Obj(a, "condition")));
        return RecomputeView(sheet, st);
    }

    object ClearFilters(JsonElement a) { Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); st.Engine.ClearConditions(); return RecomputeView(sheet, st); }

    object SetSort(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet);
        var keys = new List<(int, bool)>();
        if (a.TryGetProperty("keys", out var arr) && arr.ValueKind == JsonValueKind.Array)
            foreach (var k in arr.EnumerateArray()) { int col = Arg.Int(k, "col"); if (col < 0 || col >= ColsOf(sheet)) throw new UserError("That column does not exist."); keys.Add((col, Arg.Bool(k, "desc"))); }
        st.Sort = keys; st.Engine.SetSort(keys);
        return RecomputeView(sheet, st);
    }

    object SetHeaderRow(JsonElement a)
    {
        Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); st.Header = Arg.Bool(a, "hasHeader", true);
        return RecomputeView(sheet, st);
    }

    // =====================================================================================
    // Find and replace
    // =====================================================================================
    object Find(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet", -1); string query = Arg.Str(a, "query") ?? ""; bool matchCase = Arg.Bool(a, "matchCase");
        findCts?.Cancel(); var cts = findCts = new CancellationTokenSource();
        var which = sheet < 0 ? Enumerable.Range(0, c.Sheets.Count).ToList() : new List<int> { St(sheet) is { } && true ? sheet : 0 };
        var cacheNow = c; var cellsNow = cells!;
        var overlay = Snapshot();
        Task.Run(() =>
        {
            var batch = new List<object[]>(); int total = 0; var found = new HashSet<(int, int, int)>(); var rds = new Dictionary<int, SheetReader>(); bool completed = true;
            void Flush() { if (batch.Count == 0) return; var b = batch; batch = new(); Emit("findHits", new { Hits = b }); }
            try
            {
                void Hit(int sheetIndex, int row, int col) {
                    if (total >= SearchEngine.MaxHits) { completed = false; return; }
                    if (!rds.TryGetValue(sheetIndex, out var rd)) rds[sheetIndex] = rd = new SheetReader(cacheNow, cacheNow.Sheets[sheetIndex]);
                    string text = cellsNow.Render(sheetIndex, rd, row, col).Text;
                    if (!text.Contains(query, matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) || !found.Add((sheetIndex, row, col))) return;
                    total++; batch.Add(new object[] { sheetIndex, row, col, text });
                    if (batch.Count >= 200) Flush();
                }
                bool cacheComplete = SearchEngine.Run(cacheNow, which, query, h => { var (lr, lc) = cellsNow.Logical(h.Sheet, h.Row, h.Col); if (lr >= 0) Hit(h.Sheet, lr, lc); },
                    p => Emit("findProgress", new { Percent = p }), cts.Token);
                completed &= cacheComplete;
                foreach (var (key, edit) in overlay) {
                    cts.Token.ThrowIfCancellationRequested();
                    if (which.Contains(key.Sheet)) Hit(key.Sheet, key.Row, key.Col);
                }

            }
            catch (OperationCanceledException) { return; }
            finally { foreach (var r in rds.Values) r.Dispose(); }
            Flush(); Emit("findDone", new { Total = total, Capped = !completed });
        });
        return new { started = true };
    }

    object Replace(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        string query = Arg.Str(a, "query") ?? "", repl = Arg.Str(a, "replacement") ?? ""; bool all = Arg.Bool(a, "all"), mc = Arg.Bool(a, "matchCase");
        if (query.Length == 0) throw new UserError("Type what to find first.");
        var cmp = mc ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var list = new List<(int, int, CellEdit?)>(); int count = 0;
        void Try(SheetReader rd, int real, int col)
        {
            string? t; bool edited = cells!.TryEdit(sheet, real, col, out var e) && !EditMarks.IsStyleOnly(e);
            if (edited) t = e!.Kind == EditKind.Text ? e.Text : null; else t = cells.FileValue(sheet, rd, real, col).Text;
            if (string.IsNullOrEmpty(t) || !t.Contains(query, cmp)) return;
            var fresh = (edited ? e! : new CellEdit(EditKind.Text)) with { Kind = EditKind.Text, Text = t.Replace(query, repl, cmp) };
            list.Add((real, col, KeepStyle(sheet, real, col, fresh))); count++;
            if (list.Count > MaxCells) throw new UserError($"Too many cells would change. The limit is {MaxCells:N0}.");
        }
        lock (readerLock)
        {
            var rd = Reader(sheet);
            if (!all)
            {
                if (!a.TryGetProperty("at", out var at) || at.GetArrayLength() < 2) throw new UserError("Choose a cell first.");
                int rr = at[0].GetInt32(), cc = at[1].GetInt32();
                if (rr < 0 || rr >= RowsOf(sheet) || cc < 0 || cc >= ColsOf(sheet)) throw new UserError("That cell is outside the sheet.");
                Try(rd, rr, cc);
            }
            else
            {
                for (int ci = 0; ci < d.Chunks.Count; ci++)
                {
                    int n = rd.RowsInChunk(ci), r0 = ci * d.BlockRows;
                    for (int col = 0; col < d.Cols; col++)
                    {
                        var ch = rd.LoadColumn(ci, col); if (ch.Blob == null) continue;
                        for (int i = 0; i < n; i++) if (ch.KindAt(i) == CellKind.Text) { var (lr, lc) = cells!.Logical(sheet, r0 + i, col); if (lr >= 0) Try(rd, lr, lc); }
                        if (list.Count > MaxCells) throw new UserError($"Too many cells would change. The limit is {MaxCells:N0}.");
                    }
                }
                // Edited text cells that are not in the file's text (typed over empty or numbers).
                foreach (var (key, e) in Snapshot()) if (key.Sheet == sheet && e.Kind == EditKind.Text && !list.Any(x => x.Item1 == key.Row && x.Item2 == key.Col)) Try(rd, key.Row, key.Col);
            }
        }
        if (count == 0) return new { sheet, changed = new List<int[]>(), count = 0, dirty = Dirty, canUndo = Edits.CanUndo, canRedo = Edits.CanRedo };
        var res = Mutate(sheet, st, list);
        var t2 = res.GetType(); var changed = t2.GetProperty("changed")!.GetValue(res);
        return new { sheet, changed, count, dirty = Dirty, canUndo = Edits.CanUndo, canRedo = Edits.CanRedo };
    }

    List<KeyValuePair<(int Sheet, int Row, int Col), CellEdit>> Snapshot() { lock (editLock) return Edits.All.ToList(); }

    // =====================================================================================
    // Statistics, edge, locate
    // =====================================================================================
    object Stats(JsonElement a)
    {
        gate.Wait();
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        var ranges = Arg.Ranges(a); int vrows = ViewRows(c, st, sheet); long area = 0;
        var rects = ranges.Select(r => (R1: Math.Max(0, r.R1), C1: Math.Max(0, r.C1), R2: Math.Min(vrows - 1, r.R2), C2: Math.Min(ColsOf(sheet) - 1, r.C2))).Where(r => r.R1 <= r.R2 && r.C1 <= r.C2).ToList();
        foreach (var r in rects) area += (long)(r.R2 - r.R1 + 1) * (r.C2 - r.C1 + 1);
        if (area > 2_000_000) return new { count = area, numbers = (long?)null, cheap = true };
        long count = 0, numbers = 0; double sum = 0, min = double.MaxValue, max = double.MinValue;
        lock (readerLock)
        {
            var rd = Reader(sheet);
            foreach (var r in rects)
                for (int vr = r.R1; vr <= r.R2; vr++)
                {
                    int real = Real(st, vr);
                    for (int col = r.C1; col <= r.C2; col++)
                    {
                        byte kind; double num = 0;
                        if (cells!.TryEdit(sheet, real, col, out var e) && !EditMarks.IsStyleOnly(e))
                        {
                            if (e.Kind == EditKind.Formula) (kind, num) = cells.FormulaValue(sheet, real, col, e.Formula ?? "");
                            else { kind = e.Kind switch { EditKind.Clear => (byte)0, EditKind.Number => (byte)1, EditKind.Bool => (byte)3, _ => (byte)(string.IsNullOrEmpty(e.Text) ? 0 : 2) }; num = e.Num; }
                        }
                        else
                        {
                            var fv = cells.FileValue(sheet, rd, real, col); kind = fv.Kind;
                            if (kind == CellKind.Number) num = fv.Num;
                            if (kind == CellKind.Number && num == 0 && d.HideZeros) kind = 0;
                        }
                        if (kind == 0) continue;
                        count++;
                        if (kind == CellKind.Number) { numbers++; sum += num; if (num < min) min = num; if (num > max) max = num; }
                    }
                }
        }
        return new { count, numbers, sum, average = numbers > 0 ? sum / numbers : (double?)null, min = numbers > 0 ? min : (double?)null, max = numbers > 0 ? max : (double?)null };
    }

    bool IsEmptyCell(int sheet, SheetReader rd, int real, int col)
    {
        if (cells!.TryEdit(sheet, real, col, out var e) && !EditMarks.IsStyleOnly(e))
            return e.Kind == EditKind.Clear || e.Kind == EditKind.Text && string.IsNullOrEmpty(e.Text);
        return cells.FileValue(sheet, rd, real, col).Kind == CellKind.Empty;
    }

    object Edge(JsonElement a)
    {
        gate.Wait();
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        // Rows in and out are real sheet rows (as findHits and locate use); the walk follows the shown rows.
        int realIn = Arg.Int(a, "row"), col = Arg.Int(a, "col"); int row = realIn >= 0 && realIn < RowsOf(sheet) ? ToView(sheet, realIn) : -1;
        if (row < 0) throw new UserError("That cell is not shown.");
        CheckCell(c, st, sheet, row, col);
        string dir = Arg.Str(a, "direction") ?? "down";
        int dr = dir == "down" ? 1 : dir == "up" ? -1 : 0, dc = dir == "right" ? 1 : dir == "left" ? -1 : 0;
        int vrows = ViewRows(c, st, sheet), cols = ColsOf(sheet);
        lock (readerLock)
        {
            var rd = Reader(sheet);
            bool Empty(int r, int k) => IsEmptyCell(sheet, rd, Real(st, r), k);
            bool In(int r, int k) => r >= 0 && r < vrows && k >= 0 && k < cols;
            int r0 = row, k0 = col, rn = r0 + dr, kn = k0 + dc;
            if (!In(rn, kn)) return new { row = realIn, col };
            if (!Empty(r0, k0) && !Empty(rn, kn))
            {
                while (In(rn + dr, kn + dc) && !Empty(rn + dr, kn + dc)) { rn += dr; kn += dc; }
                return new { row = Real(st, rn), col = kn };
            }
            while (In(rn, kn) && Empty(rn, kn)) { rn += dr; kn += dc; }
            if (!In(rn, kn)) { rn -= dr; kn -= dc; }
            return new { row = Real(st, rn), col = kn };
        }
    }

    object Locate(JsonElement a)
    {
        gate.Wait();
        Need(); int sheet = Arg.Int(a, "sheet"); St(sheet); int row = Arg.Int(a, "row");
        return new { view = ToView(sheet, row) };
    }

    // =====================================================================================
    // Saving
    // =====================================================================================
    async Task<object?> SaveAsync(bool saveAs)
    {
        var c = Need(); string source = path!;
        if (Busy) throw new UserError("Please wait until the current work is finished.");
        string ext = System.IO.Path.GetExtension(source).ToLowerInvariant();
        string target = source;
        if (saveAs)
        {
            string? chosen = PickSavePath(source);
            if (chosen == null) return new { cancelled = true };
            target = System.IO.Path.GetFullPath(chosen);
            string te = System.IO.Path.GetExtension(target).ToLowerInvariant();
            if (te is not (".xlsx" or ".xlsm" or ".xlsb" or ".xls")) throw new UserError("Save this as an Excel workbook (.xlsx, .xlsm or .xlsb).");
        }
        else
        {
            if (ext is ".csv" or ".xltx" or ".xltm") throw new UserError("Save this as an Excel workbook instead. Use Save As.");
            if (!c.IsFreshFor(source)) throw new UserError("The file changed on disk since you opened it. Use Save As, or read the file again and redo your changes.");
        }
        double commit = Machine.Commit();
        if (commit >= Machine.StopAt) throw new UserError($"Windows memory is at {commit:F0}%. Saving now could hang the PC. Close some programs and try again.");
        bool same = string.Equals(System.IO.Path.GetFullPath(source), target, StringComparison.OrdinalIgnoreCase);
        if (same && !c.IsFreshFor(source)) throw new UserError("The file changed on disk since you opened it. Use Save As, or read the file again and redo your changes.");
        string job = BuildJob(c, source, target, same);
        int keepSheet = current; var sw = Stopwatch.StartNew(); bool sawSaved = false; int seenCode = -1;
        saving = true; saveOpened = false; gate.Reset(); ReleaseReaders(); watcher?.Dispose(); watcher = null;
        try
        {
            Emit("saving", new { Phase = "Starting Excel", Percent = 0, CanCancel = true });
            var ch = StartChild("--save", null, line =>
            {
                if (line == "OPENING") Emit("saving", new { Phase = "Starting Excel", Percent = 0, CanCancel = true });
                else if (line == "OPENED") { saveOpened = true; Emit("saving", new { Phase = "Writing changes", Percent = 0 }); }
                else if (line.StartsWith("WRITING ") && int.TryParse(line.AsSpan(8), out int p)) Emit("saving", new { Phase = "Writing changes", Percent = p });
                else if (line == "SAVING") Emit("saving", new { Phase = "Saving", Percent = 100 });
                else if (line == "SAVED") { sawSaved = true; Emit("saving", new { Phase = "Reading the saved file back", Percent = 0 }); }
                else if (line.StartsWith("READ ") && int.TryParse(line.AsSpan(5), out int q)) Emit("saving", new { Phase = "Reading the saved file back", Percent = q });
                else if (line == "CHECKING") Emit("saving", new { Phase = "Checking the saved file", Percent = 100 });
                else if (line.StartsWith("RESULT code=")) { var m = line.Split(' ')[1]; int.TryParse(m.AsSpan(5), out seenCode); }
            });
            child = ch; ch.P.StandardInput.WriteLine(job); ch.P.StandardInput.Flush();
            int code = await WaitChild(ch, true); child = null;
            if (code == 0)
            {
                Journal?.Forget(source); recovery = null;
                LoadCache(target, CachePolicy.PathFor(target), keepSheet); AddRecent(target);
                Log.Write($"saved edits ok {sw.ElapsedMilliseconds} ms");
                Emit("saved", new { Path = target, Seconds = Math.Round(sw.Elapsed.TotalSeconds, 1) });
                Emit("opened", new { Workbook = WorkbookInfo(), Keep = true });
                return new { saved = true, path = target };
            }
            if (sawSaved)
            {
                // The file is saved but the read-back failed: the edits are in the file now.
                Journal?.Forget(source); recovery = null;
                lock (editLock) Edits.Clear(); path = target;
                Emit("saved", new { Path = target, Seconds = Math.Round(sw.Elapsed.TotalSeconds, 1) });
                opening = true; _ = Task.Run(() => OpenCore(target));
                return new { saved = true, path = target };
            }
            string msg = code switch
            {
                4 => "The file is open in another program or is read-only. Close it there and save again, or use Save As.",
                6 => "A sheet is protected, so its cells cannot be changed.",
                7 => "The file changed on disk since you opened it. Use Save As to choose a different file.",
                2 => "Saving was cancelled. Your changes are still here.",
                90 => $"Stopped because Windows memory reached {Machine.StopAt:F0}%. Your changes are still here. Close some programs and try again.",
                _ => "Excel could not save the file. Your changes are still here.",
            };
            Emit("saveFailed", new { Error = msg });
            throw new UserError(msg);
        }
        catch (UserError) { throw; }
        catch (Exception e) when (e is IOException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            Emit("saveFailed", new { Error = "Saving could not start." }); throw new UserError("Saving could not start.");
        }
        finally
        {
            saving = false; child = null; gate.Set();
            if (cache != null && path != null && watcher == null) Watch(path);
        }
    }

    string BuildJob(CacheData c, string source, string target, bool same)
    {
        var edits = new List<object>(); var widths = new List<object>(); var freeze = new List<object>(); var structure = new List<object>();
        lock (editLock)
        {
            foreach (var op in Edits.Ops)
            {
                var d = c.Sheets[op.Sheet];
                structure.Add(new { sheet = d.Name, rows = op.Rows, at = (op.Rows ? d.FirstRow : d.FirstCol) + op.At, count = op.Count, insert = op.Insert });
            }
            foreach (var (key, e) in Edits.All)
            {
                var d = c.Sheets[key.Sheet]; bool so = EditMarks.IsStyleOnly(e);
                edits.Add(new
                {
                    sheet = d.Name, row = d.FirstRow + key.Row, col = d.FirstCol + key.Col, kind = e.Kind.ToString(), num = e.Num,
                    text = so ? null : e.Text, formula = e.Formula, @bool = e.Bool, style = e.Style, styleOnly = so,
                });
            }
            foreach (var kv in Edits.ColWidths) { var d = c.Sheets[kv.Key.Sheet]; widths.Add(new { sheet = d.Name, col = d.FirstCol + kv.Key.Col, px = kv.Value }); }
            foreach (var kv in Edits.Freeze) freeze.Add(new { sheet = c.Sheets[kv.Key].Name, rows = kv.Value.Rows, cols = kv.Value.Cols });
        }
        return JsonSerializer.Serialize(new { source, target, sameFile = same, expectedLength = same ? (long?)c.SourceLength : null, expectedTicks = same ? (long?)c.SourceTicks : null, edits, colWidths = widths, freeze, structure },
            new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
    }
}
