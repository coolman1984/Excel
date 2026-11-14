// Host tests: the message bridge, session, edits, filters, find, save orchestration, and the policy scan.
// Live tests start a hidden Excel of their own on made-up workbooks in D:\ExcelAppTest.
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using ExcelApp;
using ExcelViewer.Core;

public static class HostTests
{
    const string Dir = @"D:\ExcelAppTest";
    static Action<string, bool, string> check = (_, _, _) => { };
    static int nextId = 1;
    static readonly List<string> events = new();

    // ---- helpers -----------------------------------------------------------------------------
    static JsonElement Call(Bridge b, string cmd, object? args = null, bool expectOk = true)
    {
        var d = new Dictionary<string, object?> { ["id"] = nextId++, ["cmd"] = cmd };
        if (args != null) foreach (var p in args.GetType().GetProperties()) d[p.Name] = p.GetValue(args);
        string reply = b.Dispatch(JsonSerializer.Serialize(d)).GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(reply); var root = doc.RootElement.Clone();
        if (root.GetProperty("ok").GetBoolean() != expectOk) throw new InvalidOperationException($"{cmd} gave ok={!expectOk}" + (expectOk ? "" : "") + (root.TryGetProperty("error", out var er) ? ": " + er.GetString() : ""));
        return expectOk ? root.GetProperty("result") : root.GetProperty("error");
    }

    static JsonElement Event(string name, double seconds)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        while (sw.Elapsed.TotalSeconds < seconds)
        {
            lock (events)
                for (int i = 0; i < events.Count; i++)
                {
                    using var doc = JsonDocument.Parse(events[i]);
                    if (doc.RootElement.GetProperty("event").GetString() == name) { var r = doc.RootElement.Clone(); events.RemoveAt(i); return r; }
                }
            Thread.Sleep(50);
        }
        throw new TimeoutException(name);
    }

    static JsonElement Cell(Bridge b, int sheet, int row, int col)
    {
        var r = Call(b, "rows", new { sheet, from = row, count = 1 });
        return r.GetProperty("rows")[0][1][col];
    }
    static string Text(Bridge b, int sheet, int row, int col) => Cell(b, sheet, row, col)[0].GetString()!;
    static int Kind(Bridge b, int sheet, int row, int col) => Cell(b, sheet, row, col)[1].GetInt32();
    static int ViewRows(JsonElement r) => r.GetProperty("viewRows").GetInt32();

    static void Try(string name, Action a)
    {
        try { a(); }
        catch (Exception e) { check(name, false, e.GetType().Name + (e is InvalidOperationException ? " " + e.Message : "")); }
    }

    // ---- entry ---------------------------------------------------------------------------------
    public static void Run(Action<string, bool, string> c, bool offline)
    {
        check = c;
        CultureInfo.DefaultThreadCurrentCulture = new CultureInfo("en-US"); CultureInfo.CurrentCulture = new CultureInfo("en-US");
        Log.Folder = Path.Combine(Path.GetTempPath(), "ExcelAppTestLog");
        Offline();
        if (offline) return;
        Live();
    }

    // ---- tests that need no Excel ------------------------------------------------------------
    static void Offline()
    {
        check("jump: capped samples remain TIMEOUT in percentiles and worst",
            JumpSamples.Summary(new double?[] { 20, null, 10 }) == "JUMP median 20 ms, p95 TIMEOUT, worst TIMEOUT; TIMEOUT 1/3", "");
        check("jump: completed samples keep their measured latency",
            JumpSamples.Summary(new double?[] { 30, 10, 20 }) == "JUMP median 20 ms, p95 30 ms, worst 30 ms; TIMEOUT 0/3", "");

        Try("clipboard text splitting", () =>
        {
            var g = Session.SplitClipboard("a\tb\r\nc\t\"x\ty\"\r\n\"say \"\"hi\"\"\"\td\n");
            check("split: three rows", g.Count == 3, g.Count.ToString());
            check("split: plain cells", g[0].SequenceEqual(new[] { "a", "b" }), string.Join("|", g[0]));
            check("split: a quoted cell keeps its tab", g[1].SequenceEqual(new[] { "c", "x\ty" }), string.Join("|", g[1]));
            check("split: doubled quotes are one quote", g[2].SequenceEqual(new[] { "say \"hi\"", "d" }), string.Join("|", g[2]));
            var inverse = Session.SplitClipboard(CopyText.Build(new[] { (0, 0), (0, 1) }, (r, k) => k == 0 ? "p\tq" : "z"));
            check("split: the inverse of the copy text", inverse.Count == 1 && inverse[0][0] == "p\tq" && inverse[0][1] == "z", "");
            check("split: empty text gives no rows", Session.SplitClipboard("").Count == 0, "");
            bool capped = false;
            try { Session.SplitClipboard(new string('\t', Session.MaxCells)); } catch (UserError) { capped = true; }
            check("review: clipboard parser stops at its cell cap", capped, "");
        });
        Try("style json", () =>
        {
            var j = Cells.StyleJson(new StyleDelta(Bold: true, Fill: 0xFFFF00, Align: 1));
            check("style json: colours are #RRGGBB", j != null && (string)j["fill"] == "#FFFF00" && (bool)j["bold"] && !j.ContainsKey("italic"), "");
            check("style json: empty is null", Cells.StyleJson(null) == null && Cells.StyleJson(new StyleDelta()) == null, "");
        });
        Try("bridge errors", () =>
        {
            using var s = new Session(Path.Combine(Path.GetTempPath(), "ExcelAppTestData")); var b = new Bridge(s);
            check("bridge: unknown command is an error", Call(b, "nope", null, false).GetString()!.Contains("not known"), "");
            check("bridge: no file open is an error", Call(b, "workbook", null, false).GetString() == "No file is open.", "");
            using var bad = JsonDocument.Parse(b.Dispatch("{ not json").GetAwaiter().GetResult());
            check("bridge: broken JSON gives a plain error", !bad.RootElement.GetProperty("ok").GetBoolean(), "");
            check("bridge: error text carries no exception message", !bad.RootElement.GetProperty("error").GetString()!.Contains("Unexpected"), "");
            Call(b, "log", new { text = new string('x', 300) });
            var sentinel = "syn-log-" + Guid.NewGuid().ToString("N");
            Call(b, "log", new { text = sentinel }); Call(b, sentinel, null, false);
            check("review: arbitrary UI text and unknown commands do not reach the log", !File.ReadAllText(Log.FilePath).Contains(sentinel), "");
            check("bridge: ready reports the version", Call(b, "ready").GetProperty("version").GetString() == Session.Version, "");
            bool got = false; s.CloseAnswer = a => got = a == "dontsave"; Call(b, "closeAnswer", new { action = "dontsave" });
            check("bridge: closeAnswer reaches the window", got, "");
        });
        Try("save refuses at the memory stop", () =>
        {
            var o = new StringWriter(); int code = SaveBack.Run(new StringReader("{}"), o, 0);
            check("save: a stop at 0% refuses before starting Excel", code == 90 && o.ToString().Contains("STOPPED memory"), $"code={code}");
        });
        Try("review build memory stop", () => check("review: build refuses at memory stop before starting Excel", WorkbookReader.Build(@"D:\ExcelAppTest\syn-missing.xlsx", _ => { }, () => false, 0).Code == 90, ""));
        Try("policy scan", Policy);
    }

    static string Root()
    {
        string? env = Environment.GetEnvironmentVariable("EXCELAPP_ROOT"); if (env != null) return env;
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d != null && !Directory.Exists(Path.Combine(d.FullName, "ExcelViewer.Tests"))) d = d.Parent;
        return d?.FullName ?? throw new DirectoryNotFoundException();
    }

    // Plan step 8.11: only the files that should be able to write or save may contain the calls.
    static void Policy()
    {
        string root = Root();
        string[] forbidden = { "SaveAs(", ".Save()", "FileMode.Create", "File.WriteAll", "File.Copy", "File.Move", "File.Delete", "HttpClient", "WebClient", "Socket" };
        var bad = new List<string>();
        foreach (var f in Directory.GetFiles(Path.Combine(root, "ExcelApp"), "*.cs", SearchOption.TopDirectoryOnly))
        {
            string name = Path.GetFileName(f);
            if (name is "SaveBack.cs" or "Log.cs") continue; // Excel's own save; the log file under %LOCALAPPDATA%\ExcelApp
            int n = 0;
            foreach (var line in File.ReadAllLines(f))
            {
                n++; var code = line.Trim(); if (code.StartsWith("//")) continue;
                foreach (var w in forbidden)
                    if (code.Contains(w) && !(name == "Session.cs" && code.Contains("RecentFile")) && !(name == "Program.cs" && code.Contains("png"))) bad.Add($"{name}:{n} {w}");
            }
        }
        check("policy: no write, copy, delete or network call outside the allowed files", bad.Count == 0, string.Join("; ", bad));
        var banned = new List<string>();
        foreach (var dir in new[] { "ExcelApp", "ExcelViewer" })
            foreach (var f in Directory.EnumerateFiles(Path.Combine(root, dir), "*.cs", SearchOption.AllDirectories))
            {
                if (f.Contains(@"\obj\") || f.Contains(@"\bin\") || f.Contains(@"\ui\") || f.Contains(@"\app\")) continue;
                string t = File.ReadAllText(f);
                foreach (var w in new[] { "SheetJS", "ClosedXML", "OpenXml", "EPPlus" }) if (t.Contains(w)) banned.Add(Path.GetFileName(f) + " " + w);
            }
        string pkg = Path.Combine(root, "ExcelApp", "ui", "package.json");
        if (File.Exists(pkg)) { string t = File.ReadAllText(pkg); foreach (var w in new[] { "\"xlsx\"", "sheetjs", "exceljs", "\"xlsx-" }) if (t.Contains(w, StringComparison.OrdinalIgnoreCase)) banned.Add("package.json " + w); }
        check("policy: no workbook parser or writer library is named", banned.Count == 0, string.Join("; ", banned));
    }

    // ---- live tests (hidden Excel, made-up data) ---------------------------------------------
    static void MakeWorkbook(string path, bool password = false)
    {
        dynamic xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!; dynamic wb = null!;
        try
        {
            xl.Visible = false; xl.DisplayAlerts = false; xl.AutomationSecurity = 3;
            wb = xl.Workbooks.Add();
            while ((int)wb.Worksheets.Count < 2) wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            dynamic s1 = wb.Worksheets[1], s2 = wb.Worksheets[2]; s1.Name = "S1"; s2.Name = "S2";
            var a = new object[22, 5];
            a[0, 0] = "Name"; a[0, 1] = "Qty"; a[0, 2] = "Price"; a[0, 3] = "Note"; a[0, 4] = "Tail";
            for (int n = 1; n <= 20; n++) { a[n, 0] = n == 4 ? "needle-unique" : "item-" + n; a[n, 1] = (double)n; a[n, 2] = n * 1.5; }
            a[21, 0] = "kinds"; a[21, 1] = 3.5; a[21, 2] = true;
            s1.Range["A1:E22"].Value2 = a;
            s1.Range["C2:C21"].NumberFormat = "$#,##0.00";
            s1.Range["D22"].Formula = "=1/0";
            s2.Range["A1"].Value2 = "x"; s2.Range["A2"].Value2 = 5.0;
            if (File.Exists(path)) File.Delete(path);
            if (password) wb.SaveAs(path, 51, Password: "syn-password"); else wb.SaveAs(path, 51);
        }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch (COMException) { }
            try { xl.Quit(); } catch (COMException) { }
            if (wb != null && Marshal.IsComObject(wb)) Marshal.FinalReleaseComObject(wb);
            Marshal.FinalReleaseComObject(xl); GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }

    static string FindChildExe()
    {
        string here = Path.Combine(AppContext.BaseDirectory, "ExcelApp.exe");
        if (File.Exists(here)) return here;
        return Path.Combine(Root(), "ExcelApp", "bin", "Release", "net8.0-windows", "ExcelApp.exe");
    }

    static void Live()
    {
        Directory.CreateDirectory(Dir);
        string file = Path.Combine(Dir, "syn-host.xlsx"), saved = Path.Combine(Dir, "syn-host-saved.xlsx");
        if (!CachePolicy.ThisPcApproved) { check("live: this PC is approved for the fast copy", false, ""); return; }
        MakeWorkbook(file);
        Try("review password refusal", () => {
            string encrypted = Path.Combine(Dir, "syn-password.xlsx");
            MakeWorkbook(encrypted, true);
            try {
                var timer = System.Diagnostics.Stopwatch.StartNew();
                var result = WorkbookReader.Build(encrypted, _ => { }, () => false);
                check("review: password workbook refused without a prompt", result.Code == 1 && timer.Elapsed.TotalSeconds < 30, "");
            } finally { if (File.Exists(encrypted)) File.Delete(encrypted); }
        });
        CachePolicy.Delete(CachePolicy.PathFor(file)); CachePolicy.Delete(CachePolicy.PathFor(saved));
        if (File.Exists(saved)) File.Delete(saved);
        using var s = new Session(Path.Combine(Path.GetTempPath(), "ExcelAppTestData")) { ChildExe = FindChildExe() };
        var b = new Bridge(s); b.Post = j => { lock (events) events.Add(j); };
        var clip = new List<string>(); string clipText = "";
        s.ClipboardSet = t => { clip.Add(t); clipText = t; }; s.ClipboardGet = () => clipText;

        bool opened = false;
        Try("open builds the fast copy and reports progress", () =>
        {
            Call(b, "open", new { path = file });
            var ev = Event("opened", 300); opened = true;
            check("open: the opened event carries the workbook", ev.GetProperty("workbook").GetProperty("sheets").GetArrayLength() == 2, "");
        });
        if (!opened) return;
        var wb = Call(b, "workbook");
        var sh0 = wb.GetProperty("sheets")[0];
        Try("workbook shape", () =>
        {
            check("workbook: rows and columns", sh0.GetProperty("rows").GetInt32() == 22 && sh0.GetProperty("cols").GetInt32() == 5, $"{sh0.GetProperty("rows")}x{sh0.GetProperty("cols")}");
            check("workbook: first row is 1", sh0.GetProperty("firstRow").GetInt32() == 1 && sh0.GetProperty("colWidthsPx").GetArrayLength() == 5, "");
            check("workbook: a file with no sidecar is clean", !wb.GetProperty("dirty").GetBoolean(), "");
            check("sheet: selecting reports the view", ViewRows(Call(b, "sheet", new { sheet = 0 })) == 22, "");
        });
        Try("rows wire format", () =>
        {
            check("rows: text cell", Text(b, 0, 1, 0) == "item-1" && Kind(b, 0, 1, 0) == 2, "");
            var n = Cell(b, 0, 1, 1); check("rows: number is kind 1, right aligned", n[0].GetString() == "1" && n[1].GetInt32() == 1 && n[2].GetInt32() == 1, "");
            check("rows: money keeps its format", Text(b, 0, 1, 2) == "$1.50", Text(b, 0, 1, 2));
            check("rows: every kind in one row", Kind(b, 0, 21, 0) == 2 && Kind(b, 0, 21, 1) == 1 && Kind(b, 0, 21, 2) == 3 && Kind(b, 0, 21, 3) == 4 && Kind(b, 0, 21, 4) == 0, "");
            check("rows: bool and error text", Text(b, 0, 21, 2) == "TRUE" && Text(b, 0, 21, 3) == "#DIV/0!", "");
            var r = Call(b, "rows", new { sheet = 0, from = 0, count = 3 });
            check("rows: original Excel row numbers", r.GetProperty("rows")[0][0].GetInt32() == 1 && r.GetProperty("rows")[2][0].GetInt32() == 3, "");
            check("rows: at most the rows that exist", Call(b, "rows", new { sheet = 0, from = 20, count = 500 }).GetProperty("rows").GetArrayLength() == 2, "");
            var d = Call(b, "cellDetail", new { sheet = 0, row = 2, col = 1 });
            check("cellDetail: address and raw value", d.GetProperty("address").GetString() == "B3" && d.GetProperty("raw").GetString() == "2" && !d.GetProperty("edited").GetBoolean(), "");
        });
        Try("typed input", () =>
        {
            var r = Call(b, "setCell", new { sheet = 0, row = 1, col = 3, input = "12.5%" });
            check("setCell: reports the changed cell and dirty", r.GetProperty("dirty").GetBoolean() && r.GetProperty("changed").GetArrayLength() == 1, "");
            check("setCell: 12.5% in a General cell shows 12.50%", Text(b, 0, 1, 3) == "12.50%", Text(b, 0, 1, 3));
            check("setCell: the cell is marked edited", Call(b, "cellDetail", new { sheet = 0, row = 1, col = 3 }).GetProperty("edited").GetBoolean(), "");
            var u = Call(b, "undo", new { sheet = 0 });
            check("undo: restores the empty cell and clears dirty", Text(b, 0, 1, 3) == "" && !u.GetProperty("dirty").GetBoolean() && !u.GetProperty("canUndo").GetBoolean() && u.GetProperty("canRedo").GetBoolean(), "");
            Call(b, "redo", new { sheet = 0 }); check("redo: brings it back", Text(b, 0, 1, 3) == "12.50%", "");
            Call(b, "undo", new { sheet = 0 });
            var fr = Call(b, "setCell", new { sheet = 0, row = 1, col = 3, input = "=B2*2" });
            var c = Cell(b, 0, 1, 3);
            check("formula: shows the calculated result as a number", c[0].GetString() == "2" && c[1].GetInt32() == 1, c.ToString());
            check("formula: the wire carries the formula for the in-cell editor", c.GetArrayLength() > 6 && c[6].GetString() == "=B2*2", c.ToString());
            check("formula: the result asks the interface to repaint", fr.GetProperty("recalc").GetBoolean(), "");
            check("formula: the formula bar shows the formula", Call(b, "cellDetail", new { sheet = 0, row = 1, col = 3 }).GetProperty("raw").GetString() == "=B2*2", "");
            Call(b, "setCell", new { sheet = 0, row = 1, col = 1, input = "21" });
            check("formula: recalculates when a cell it uses changes", Text(b, 0, 1, 3) == "42", Text(b, 0, 1, 3));
            Call(b, "undo", new { sheet = 0 });
            check("formula: undo of the source recalculates too", Text(b, 0, 1, 3) == "2", Text(b, 0, 1, 3));
            Call(b, "setCell", new { sheet = 0, row = 2, col = 3, input = "=SUM(B2:B4" });
            check("formula: a missing bracket is closed, as Excel does", Call(b, "cellDetail", new { sheet = 0, row = 2, col = 3 }).GetProperty("raw").GetString() == "=SUM(B2:B4)" && Text(b, 0, 2, 3) == "6", Text(b, 0, 2, 3));
            Call(b, "setCell", new { sheet = 0, row = 3, col = 3, input = "=NOSUCH(1)" });
            check("formula: an unknown function shows #NAME?", Text(b, 0, 3, 3) == "#NAME?" && Kind(b, 0, 3, 3) == 4, Text(b, 0, 3, 3));
            Call(b, "setCell", new { sheet = 0, row = 4, col = 3, input = "=1/0" });
            check("formula: division by zero shows #DIV/0!", Text(b, 0, 4, 3) == "#DIV/0!", Text(b, 0, 4, 3));
            Call(b, "setCell", new { sheet = 0, row = 5, col = 3, input = "=SUM(1,2" + ")" + ")" }, false);
            Call(b, "setCell", new { sheet = 0, row = 5, col = 3, input = "=0.1+0.2" });
            check("formula: 15 significant digits, so 0.1+0.2 shows 0.3", Text(b, 0, 5, 3) == "0.3", Text(b, 0, 5, 3));
            var st0 = Call(b, "stats", new { sheet = 0, ranges = new[] { new[] { 1, 3, 2, 3 } } });
            check("formula: the status bar sums calculated results", Math.Abs(st0.GetProperty("sum").GetDouble() - 8) < 1e-9, st0.ToString());
            Call(b, "fill", new { sheet = 0, from = new[] { 1, 3, 1, 3 }, to = new[] { 1, 3, 3, 3 } });
            check("fill: relative references move with the formula", Call(b, "cellDetail", new { sheet = 0, row = 3, col = 3 }).GetProperty("raw").GetString() == "=B4*2" && Text(b, 0, 3, 3) == "6", Text(b, 0, 3, 3));
            Call(b, "copy", new { sheet = 0, ranges = new[] { new[] { 1, 3, 1, 3 } } });
            Call(b, "paste", new { sheet = 0, row = 2, col = 4 });
            check("paste: a copied formula is pasted as a formula, moved", Call(b, "cellDetail", new { sheet = 0, row = 2, col = 4 }).GetProperty("raw").GetString() == "=C3*2", Call(b, "cellDetail", new { sheet = 0, row = 2, col = 4 }).GetProperty("raw").GetString());
            for (int i = 0; i < 8; i++) Call(b, "undo", new { sheet = 0 });
            check("formula: undo takes every formula away", Text(b, 0, 1, 3) == "" && Text(b, 0, 3, 3) == "" && !Call(b, "workbook").GetProperty("dirty").GetBoolean(), Text(b, 0, 1, 3));
            Call(b, "setCell", new { sheet = 0, row = 1, col = 4, input = "Item 7" });
            Call(b, "fill", new { sheet = 0, from = new[] { 1, 4, 1, 4 }, to = new[] { 1, 4, 3, 4 } });
            check("fill: text ending in a number counts up", Text(b, 0, 3, 4) == "Item 9", Text(b, 0, 3, 4));
            Call(b, "undo", new { sheet = 0 }); Call(b, "undo", new { sheet = 0 });
            Call(b, "setCell", new { sheet = 0, row = 1, col = 1, input = "7" });
            check("setCell: a number over a file number", Text(b, 0, 1, 1) == "7" && Kind(b, 0, 1, 1) == 1, "");
            Call(b, "setCell", new { sheet = 0, row = 1, col = 1, input = "" });
            check("setCell: empty input clears the cell", Kind(b, 0, 1, 1) == 0, "");
            Call(b, "undo", new { sheet = 0 }); Call(b, "undo", new { sheet = 0 });
            check("undo: the file's value is back", Text(b, 0, 1, 1) == "1", "");
            Call(b, "setCell", new { sheet = 0, row = 40, col = 0, input = "x" }, false);
        });
        Try("clear, style, paste, copy, fill", () =>
        {
            Call(b, "clearCells", new { sheet = 0, ranges = new[] { new[] { 1, 0, 2, 0 } } });
            check("clearCells: contents gone", Text(b, 0, 1, 0) == "" && Text(b, 0, 2, 0) == "", "");
            Call(b, "undo", new { sheet = 0 }); check("clearCells: one undo restores both", Text(b, 0, 1, 0) == "item-1" && Text(b, 0, 2, 0) == "item-2", "");
            Call(b, "setStyle", new { sheet = 0, ranges = new[] { new[] { 1, 0, 2, 0 } }, style = new { bold = true, fill = "#FFFF00" } });
            var st = Cell(b, 0, 1, 0)[5];
            check("setStyle: wire style has bold and fill, value untouched", st.GetProperty("bold").GetBoolean() && st.GetProperty("fill").GetString() == "#FFFF00" && Text(b, 0, 1, 0) == "item-1", "");
            var d = Call(b, "cellDetail", new { sheet = 0, row = 1, col = 0 });
            check("cellDetail: reports the style", d.GetProperty("style").GetProperty("bold").GetBoolean(), "");
            Call(b, "setStyle", new { sheet = 0, ranges = new[] { new[] { 1, 1, 1, 1 } }, style = new { align = 2, numberFormat = "0.00%" } });
            check("setStyle: a number format re-renders the number", Text(b, 0, 1, 1) == "100.00%", Text(b, 0, 1, 1));
            Call(b, "undo", new { sheet = 0 }); Call(b, "undo", new { sheet = 0 });
            check("undo: style edits go too", Cell(b, 0, 1, 0).GetArrayLength() <= 5 || Cell(b, 0, 1, 0)[5].ValueKind == JsonValueKind.Null, "");

            clipText = "a\t1\r\nb\t2\r\nc\t3";
            var p = Call(b, "paste", new { sheet = 0, row = 5, col = 0 });
            check("paste: a 3 x 2 block makes 6 changes", p.GetProperty("changed").GetArrayLength() == 6, p.GetProperty("changed").GetArrayLength().ToString());
            check("paste: values landed", Text(b, 0, 5, 0) == "a" && Text(b, 0, 7, 1) == "3", "");
            Call(b, "undo", new { sheet = 0 });
            check("paste: one undo removes the whole paste", Text(b, 0, 5, 0) == "item-5" && Text(b, 0, 7, 1) == "7", "");
            Call(b, "redo", new { sheet = 0 }); Call(b, "undo", new { sheet = 0 });

            var cp = Call(b, "copy", new { sheet = 0, ranges = new[] { new[] { 0, 0, 1, 1 } } });
            check("copy: shown text with tabs and line breaks", clip[^1] == "Name\tQty\r\nitem-1\t1" && cp.GetProperty("cells").GetInt32() == 4, clip[^1].Replace("\t", "<T>").Replace("\r\n", "<N>"));
            check("copy: over the cell limit is refused in plain words", Call(b, "copy", new { sheet = 0, ranges = new[] { new[] { 0, 0, 100000, 100 } } }, false).GetString()!.Contains("Too many"), "");

            Call(b, "cut", new { sheet = 0, ranges = new[] { new[] { 1, 0, 1, 0 } } });
            clipText = clip[^1];
            Call(b, "paste", new { sheet = 0, row = 10, col = 0 });
            check("cut and paste: the source clears and the target gets it", Text(b, 0, 1, 0) == "" && Text(b, 0, 10, 0) == "item-1", "");
            Call(b, "undo", new { sheet = 0 });
            check("cut and paste: one undo restores both", Text(b, 0, 1, 0) == "item-1" && Text(b, 0, 10, 0) == "item-10", "");

            var f = Call(b, "fill", new { sheet = 0, from = new[] { 1, 1, 1, 1 }, to = new[] { 1, 1, 4, 1 } });
            check("fill: the source repeats over the target", f.GetProperty("changed").GetArrayLength() == 3 && Text(b, 0, 4, 1) == "1", "");
            Call(b, "undo", new { sheet = 0 });
            var f2 = Call(b, "fill", new { sheet = 0, from = new[] { 1, 0, 2, 0 }, to = new[] { 1, 0, 5, 0 } });
            check("fill: two numbered text cells continue the step, as in Excel", Text(b, 0, 3, 0) == "item-3" && Text(b, 0, 4, 0) == "item-4", Text(b, 0, 3, 0) + "," + Text(b, 0, 4, 0));
            Call(b, "undo", new { sheet = 0 });
        });
        Try("agent commands", () =>
        {
            // AI agents read and write by Excel address (A1 is the header "Name", B2 holds 1 in the made-up workbook)
            var r = Call(b, "agent.read", new { range = "A1:B2" });
            check("agent: reads by Excel address", r.GetProperty("values")[0][0].GetString() == "Name" && r.GetProperty("values")[1][1].GetString() == "1", r.ToString());
            var w = Call(b, "agent.write", new { start = "B2", values = new object[][] { new object[] { 5 }, new object[] { "=B2*2" } } });
            check("agent: writes a number and a formula", w.GetProperty("written").GetInt32() == 2 && Text(b, 0, 1, 1) == "5" && Text(b, 0, 2, 1) == "10", Text(b, 0, 2, 1));
            var f = Call(b, "agent.read", new { sheet = 0, range = "B3" });
            check("agent: a read shows the formula", f.GetProperty("formulas")[0][0].GetString() == "=B2*2", f.ToString());
            Call(b, "undo", new { sheet = 0 });
            check("agent: one undo takes the whole write back", Text(b, 0, 1, 1) == "1", Text(b, 0, 1, 1));
            check("agent: outside the data is refused plainly", Call(b, "agent.write", new { start = "XFD1", values = new[] { new[] { 1 } } }, false).GetString()!.Contains("outside"), "");
            check("agent: an unknown sheet names the sheets", Call(b, "agent.read", new { sheet = "Nope", range = "A1" }, false).GetString()!.Contains("Sheets:"), "");
        });
        Try("filter and sort", () =>
        {
            var r = Call(b, "setFilter", new { sheet = 0, col = 1, condition = new { kind = "number", op = "gt", a = 10 } });
            check("filter: number greater than keeps the header and 10 rows", ViewRows(r) == 11 && r.GetProperty("filterActive").GetBoolean(), ViewRows(r).ToString());
            check("filter: header stays first, then original row numbers", Call(b, "rows", new { sheet = 0, from = 0, count = 2 }).GetProperty("rows")[0][0].GetInt32() == 1 && Call(b, "rows", new { sheet = 0, from = 0, count = 2 }).GetProperty("rows")[1][0].GetInt32() == 12, "");
            check("locate: a row hidden by the filter gives -1", Call(b, "locate", new { sheet = 0, row = 3 }).GetProperty("view").GetInt32() == -1 && Call(b, "locate", new { sheet = 0, row = 15 }).GetProperty("view").GetInt32() == 5, "");
            var sh = Call(b, "sheet", new { sheet = 0 });
            check("sheet: reports the active filter", sh.GetProperty("filters").GetProperty("1").GetProperty("kind").GetString() == "number", "");
            Call(b, "setSort", new { sheet = 0, keys = new[] { new { col = 1, desc = true } } });
            check("sort after filter keeps the header first", Call(b, "rows", new { sheet = 0, from = 0, count = 2 }).GetProperty("rows")[0][0].GetInt32() == 1 && Text(b, 0, 1, 1) == "20", Text(b, 0, 1, 1));
            r = Call(b, "clearFilters", new { sheet = 0 }); check("clearFilters: all rows back, sort stays", ViewRows(r) == 22 && Text(b, 0, 1, 1) == "20", "");
            r = Call(b, "setSort", new { sheet = 0, keys = Array.Empty<object>() });
            check("setSort: empty clears the sort", ViewRows(r) == 22 && Text(b, 0, 1, 1) == "1", "");
            r = Call(b, "setFilter", new { sheet = 0, col = 0, condition = new { kind = "values", include = new[] { "item-3" }, blanks = false } });
            check("filter: values keeps exactly the chosen text", ViewRows(r) == 2 && Text(b, 0, 1, 0) == "item-3", ViewRows(r).ToString());
            Call(b, "clearFilters", new { sheet = 0 });
            Call(b, "setFilter", new { sheet = 0, col = 1, condition = new { kind = "number", op = "gt", a = 10 } });
            r = Call(b, "setFilter", new { sheet = 0, col = 0, condition = new { kind = "text", op = "contains", value = "ITEM-1" } });
            check("filter: text contains (any case) AND a number condition", ViewRows(r) == 10, ViewRows(r).ToString());
            Call(b, "clearFilters", new { sheet = 0 });
            var fv = Call(b, "filterValues", new { sheet = 0, col = 0 });
            check("filterValues: lists the distinct texts", fv.GetProperty("values").GetArrayLength() >= 20, "");
            check("filterValues: a number column answers number filters", Call(b, "filterValues", new { sheet = 0, col = 1 }).GetProperty("kind").GetString() == "number", "");
            r = Call(b, "setHeaderRow", new { sheet = 0, hasHeader = false }); check("setHeaderRow: accepted", ViewRows(r) == 22, "");
            Call(b, "setHeaderRow", new { sheet = 0, hasHeader = true });
        });
        Try("find, replace, stats, edge", () =>
        {
            lock (events) events.Clear();
            Call(b, "find", new { sheet = -1, query = "needle-unique" });
            var done = Event("findDone", 30);
            check("find: a known needle is found once", done.GetProperty("total").GetInt32() == 1 && !done.GetProperty("capped").GetBoolean(), done.GetProperty("total").ToString());
            var hits = Event("findHits", 5).GetProperty("hits");
            check("find: the hit holds sheet, row, column and text", hits[0][0].GetInt32() == 0 && hits[0][1].GetInt32() == 4 && hits[0][2].GetInt32() == 0 && hits[0][3].GetString() == "needle-unique", "");
            var rp = Call(b, "replace", new { sheet = 0, query = "item-", replacement = "it-", all = true });
            check("replace all: changes the expected number of text cells", rp.GetProperty("count").GetInt32() == 19, rp.GetProperty("count").ToString());
            check("replace all: the text changed", Text(b, 0, 1, 0) == "it-1", "");
            Call(b, "undo", new { sheet = 0 }); check("replace all: one undo", Text(b, 0, 1, 0) == "item-1", "");
            var one = Call(b, "replace", new { sheet = 0, query = "item", replacement = "Z", all = false, at = new[] { 1, 0 } });
            check("replace one: only that cell", one.GetProperty("count").GetInt32() == 1 && Text(b, 0, 1, 0) == "Z-1" && Text(b, 0, 2, 0) == "item-2", "");
            Call(b, "undo", new { sheet = 0 });
            var st = Call(b, "stats", new { sheet = 0, ranges = new[] { new[] { 1, 1, 20, 1 } } });
            check("stats: count, sum, average, min, max", st.GetProperty("count").GetInt64() == 20 && st.GetProperty("numbers").GetInt64() == 20 && st.GetProperty("sum").GetDouble() == 210 && st.GetProperty("average").GetDouble() == 10.5 && st.GetProperty("min").GetDouble() == 1 && st.GetProperty("max").GetDouble() == 20, "");
            var e1 = Call(b, "edge", new { sheet = 0, row = 0, col = 0, direction = "down" });
            check("edge: down a full column reaches the last row", e1.GetProperty("row").GetInt32() == 21 && e1.GetProperty("col").GetInt32() == 0, "");
            var e2 = Call(b, "edge", new { sheet = 0, row = 0, col = 3, direction = "down" });
            check("edge: from the last cell before a gap, jump to the next filled cell", e2.GetProperty("row").GetInt32() == 21, e2.GetProperty("row").ToString());
            var e3 = Call(b, "edge", new { sheet = 0, row = 0, col = 0, direction = "right" });
            check("edge: right along a full row", e3.GetProperty("col").GetInt32() == 4, "");
            var e4 = Call(b, "edge", new { sheet = 0, row = 10, col = 3, direction = "up" });
            check("edge: up from an empty cell reaches the header", e4.GetProperty("row").GetInt32() == 0, e4.GetProperty("row").ToString());
        });
        Try("review regressions", () =>
        {
            Call(b, "setCell", new { sheet = 0, row = 2, col = 1, input = "overlay-needle" });
            lock (events) events.Clear();
            Call(b, "find", new { sheet = 0, query = "overlay-needle" });
            check("review: find sees text typed over a numeric cache cell", Event("findDone", 30).GetProperty("total").GetInt32() == 1, "");
            Call(b, "undo");
            Call(b, "setCell", new { sheet = 0, row = 4, col = 0, input = "gone" });
            lock (events) events.Clear(); Call(b, "find", new { sheet = 0, query = "needle-unique" });
            check("review: find excludes a cached hit overwritten by an edit", Event("findDone", 30).GetProperty("total").GetInt32() == 0, "");
            Call(b, "undo");
            Call(b, "setStyle", new { sheet = 0, ranges = new[] { new[] { 1, 1, 1, 1 } }, style = new { numberFormat = "0.00%" } });
            var fv = Call(b, "filterValues", new { sheet = 0, col = 1 });
            check("review: filters use edited number formats", fv.GetProperty("values").EnumerateArray().Any(v => v.GetProperty("text").GetString() == "100.00%"), "");
            Call(b, "undo");
            check("review: locate rejects rows beyond the sheet", Call(b, "locate", new { sheet = 0, row = 999 }).GetProperty("view").GetInt32() == -1, "");
            clipText = "bad";
            check("review: paste refuses negative coordinates", Call(b, "paste", new { sheet = 0, row = -1, col = 0 }, false).GetString()!.Contains("outside"), "");
            clipText = "item-1";
            Call(b, "cut", new { sheet = 0, ranges = new[] { new[] { 1, 0, 1, 0 } } });
            check("review: cut paste outside sheet refuses without deleting source", Call(b, "paste", new { sheet = 0, row = 999, col = 0 }, false).GetString()!.Contains("outside") && Text(b, 0, 1, 0) == "item-1", "");
            Call(b, "copy", new { sheet = 0, ranges = new[] { new[] { 1, 0, 1, 0 } } });
            var timestamp = File.GetLastWriteTimeUtc(file);
            File.SetLastWriteTimeUtc(file, timestamp.AddMinutes(1));
            s.PickSavePath = _ => file;
            check("review: Save As to the source refuses when source changed", Call(b, "save", new { saveAs = true }, false).GetString()!.Contains("changed on disk"), "");
            File.SetLastWriteTimeUtc(file, timestamp);
        });
        Try("width and freeze", () =>
        {
            Call(b, "setColWidth", new { sheet = 0, col = 0, px = 120 }); Call(b, "setFreeze", new { sheet = 0, rows = 1, cols = 1 });
            var w = Call(b, "workbook").GetProperty("sheets")[0];
            check("width and freeze are remembered and make the file dirty", w.GetProperty("colWidthsPx")[0].GetDouble() == 120 && w.GetProperty("frozenRows").GetInt32() == 1 && Call(b, "workbook").GetProperty("dirty").GetBoolean(), "");
        });
        Try("save as through Excel, then read back in the same session", () =>
        {
            Call(b, "setCell", new { sheet = 0, row = 1, col = 1, input = "25" });
            Call(b, "setCell", new { sheet = 0, row = 1, col = 3, input = "=B2*2" });
            Call(b, "setStyle", new { sheet = 0, ranges = new[] { new[] { 0, 0, 0, 0 } }, style = new { bold = true } });
            s.PickSavePath = _ => saved;
            lock (events) events.Clear();
            var r = Call(b, "save", new { saveAs = true });
            check("save: reported saved", r.GetProperty("saved").GetBoolean(), "");
            var ev = Event("saved", 5); var op = Event("opened", 5);
            check("save: the host reopened from the rebuilt copy", op.GetProperty("keep").GetBoolean() && !op.GetProperty("workbook").GetProperty("dirty").GetBoolean(), "");
            check("save: the edits are now in the file's own values", Text(b, 0, 1, 1) == "25", Text(b, 0, 1, 1));
            check("save: Excel recalculated the formula", Text(b, 0, 1, 3) == "50", Text(b, 0, 1, 3));
            check("save: the saved file is NASCA protected", File.ReadAllBytes(saved).Take(8).SequenceEqual(Encoding.ASCII.GetBytes("<## NASC")), "NASCA mark: no");
            check("save: the copy is fresh for the saved file", CacheData.HeaderMatches(CachePolicy.PathFor(saved), saved), "");
            check("save: the session now points at the saved file", s.SourcePath == saved && Call(b, "info").GetProperty("nasca").GetBoolean(), "");
            check("save: no .part file left", !File.Exists(CachePolicy.PathFor(saved) + ".part") && !File.Exists(CachePolicy.SidecarPathFor(saved) + ".part"), "");
            Call(b, "save", new { saveAs = false }); // nothing to do is fine; a second save of the same file works
            Call(b, "setCell", new { sheet = 0, row = 1, col = 1, input = "26" });
        });
        Try("clipboard flags", ClipboardFlags);
    }

    static void ClipboardFlags()
    {
        string made = "made-up clip " + Guid.NewGuid().ToString("N")[..6]; bool ok = false; string detail = "";
        var t = new Thread(() =>
        {
            try
            {
                string? before = Clipboard.ContainsText() ? Clipboard.GetText() : null;
                ExcelApp.MainForm.ClipboardSink(made);
                var data = Clipboard.GetDataObject()!;
                bool Flag(string name, int want) { using var ms = data.GetData(name) as MemoryStream; if (ms == null) return false; var bytes = ms.ToArray(); return bytes.Length >= 4 && BitConverter.ToInt32(bytes, 0) == want; }
                ok = Clipboard.GetText() == made && Flag("ExcludeClipboardContentFromMonitorProcessing", 1) && Flag("CanIncludeInClipboardHistory", 0) && Flag("CanUploadToCloudClipboard", 0);
                detail = ok ? "" : "flags missing";
                if (before != null) ExcelApp.MainForm.ClipboardSink(before); else Clipboard.Clear();
            }
            catch (Exception e) { detail = e.GetType().Name; }
        });
        t.SetApartmentState(ApartmentState.STA); t.Start(); t.Join();
        check("copy: text goes on the clipboard with the three flags", ok, detail);
    }
}
