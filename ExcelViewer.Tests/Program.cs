// Tests for the viewer. Made-up data only, no test package needed.
// Without arguments: everything, including a hidden Excel. With "offline": no Excel.
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using ExcelViewer.Core;

static class T
{
    static int pass, fail;
    static void Check(string name, bool ok, string detail = "") { if (ok) pass++; else { fail++; Console.WriteLine($"FAIL {name} {detail}"); } }

    [STAThread]
    static int Main(string[] a)
    {
        if (a.Contains("clipboard")) return RealClipboard();
        var machine = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = new CultureInfo("en-US");
        Formats(); Unsupported(); Policy(); BadCopies(); CopyPieces();
        foreach (var (x, v, ok) in new[] { ("0.00", 0.004, true), ("0", 0.004, true), ("1235", 1234.5678, true), ("1.2E+03", 1234.5678, true), ("1234.57", 1234.5678, true), ("1234.6", 1234.5678, true), ("1235", 1500.0, false), ("5", 7.0, false), ("abc", 1.0, false) })
            Check($"General agrees: {x} for {v}", CacheBuilder.GeneralAgrees(x, v) == ok);
        Console.WriteLine($"offline: {pass} passed, {fail} failed");
        CultureInfo.CurrentCulture = machine;
        if (!a.Contains("offline")) { if (Machine.Commit() >= Machine.StopAt) Console.WriteLine("SKIPPED live tests: memory at the stop line"); else Live(); }
        Console.WriteLine(fail == 0 ? $"ALL PASSED ({pass})" : $"FAILED {fail} of {pass + fail}");
        return fail == 0 ? 0 : 1;
    }

    static void F(string code, double v, string want, bool d1904 = false)
    {
        var s = new NumFmt(code, d1904).Number(v);
        Check($"format [{code}] of {v}", s != null && s.Value.Text.Trim() == want, $"got '{s?.Text}' want '{want}'");
    }

    static void Formats()
    {
        // Money
        F("#,##0.00", 1234567.891, "1,234,567.89"); F("#,##0", 999.5, "1,000"); F("#,##0.00", -1234.5, "-1,234.50"); F("0.00", 1.005, "1.01");
        F("$#,##0.00", 1234.5, "$1,234.50"); F("$#,##0.00;[Red]($#,##0.00)", -1234.5, "($1,234.50)"); F("[$EGP] #,##0.00", 1500, "EGP 1,500.00");
        F("[$€-407] #,##0.00", 12.3, "€ 12.30"); F("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", 1234.5, "1,234.50"); F("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", -1234.5, "(1,234.50)"); F("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", 0, "-");
        F("#,##0,", 1234567, "1,235"); F("0.0,,", 2500000, "2.5"); F("#,##0.00\" EGP\"", 10, "10.00 EGP");
        // Percent
        F("0%", 0.256, "26%"); F("0.00%", 0.12345, "12.35%"); F("0.0%", -0.5, "-50.0%");
        // Plain numbers
        F("0", 0.4, "0"); F("#", 0, ""); F("000-00", 12345, "123-45"); F("00000", 42, "00042"); F("#.##", 5, "5."); F("#.##", 5.1, "5.1"); F(".00", 0.5, ".50"); F(".00", 12.5, "12.50");
        F("0.00E+00", 12345, "1.23E+04"); F("0.00E+00", 0.00012, "1.20E-04"); F("# ?/?", 2.5, "2 1/2"); F("# ??/16", 2.25, "2 4/16");
        F("General", 1234.5, "1234.5"); F("General", 123456789012, "1.23457E+11"); F("General", 0.1, "0.1"); F("General", -7, "-7"); F("General\" kg\"", 5, "5 kg");
        // Sections and conditions
        F("0;-0;\"zero\"", 0, "zero"); F("0;(0)", -5, "(5)"); F("[>=100]\"big\";[<0]\"neg\";\"small\"", 150, "big"); F("[>=100]\"big\";[<0]\"neg\";\"small\"", -1, "neg"); F("[>=100]\"big\";[<0]\"neg\";\"small\"", 5, "small"); F(";;;", 5, "");
        var red = new NumFmt("#,##0;[Red]-#,##0").Number(-5); Check("red colour carried", red?.Color == System.Drawing.Color.Red);
        // Dates. 45236 is 2023-11-06, a Monday.
        F("yyyy-mm-dd", 45236, "2023-11-06"); F("dd/mm/yyyy", 45236, "06/11/2023"); F("d-mmm-yy", 45236, "6-Nov-23"); F("mmmm d, yyyy", 45236, "November 6, 2023"); F("dddd", 45236, "Monday"); F("ddd", 45236, "Mon"); F("mmmmm", 45236, "N"); F("dd.mm.yyyy", 45236, "06.11.2023");
        F("m/d/yyyy", 45236, "11/6/2023"); F("[$-409]mmm-yy", 45236, "Nov-23");
        // The 1900 quirks Excel keeps on purpose.
        F("yyyy-mm-dd", 1, "1900-01-01"); F("yyyy-mm-dd", 59, "1900-02-28"); F("yyyy-mm-dd", 60, "1900-02-29"); F("yyyy-mm-dd", 61, "1900-03-01"); F("m/d/yyyy", 0, "1/0/1900"); F("dddd", 1, "Sunday");
        F("yyyy-mm-dd", 0, "1904-01-01", true); F("yyyy-mm-dd", 45236, "2027-11-07", true);
        // Times
        F("hh:mm:ss", 0.5, "12:00:00"); F("h:mm AM/PM", 0.75, "6:00 PM"); F("h:mm AM/PM", 0, "12:00 AM"); F("h:mm am/pm", 0.75, "6:00 pm"); F("h:mm a/p", 0.25, "6:00 a"); F("[h]:mm", 1.5, "36:00"); F("[mm]:ss", 0.5, "720:00"); F("mm:ss.00", 0.0000115741, "00:01.00"); F("yyyy-mm-dd hh:mm", 45236.5, "2023-11-06 12:00");
        F("m/d/yy h:mm", 45236.99999, "11/6/23 23:59");
        // Numbers Excel cannot show as a date
        F("yyyy-mm-dd", -1, "#####"); F("yyyy-mm-dd", 2958466, "#####");
        // Text
        Check("text section", new NumFmt("0;0;0;\"[\"@\"]\"").Text("abc") == "[abc]"); Check("text format", new NumFmt("@").Text("abc") == "abc"); Check("text with number format", new NumFmt("0.00").Text("abc") == "abc"); Check("hide text", new NumFmt(";;;").Text("abc") == "");
    }

    static void Unsupported()
    {
        foreach (var code in new[] { "[DBNum1]0", "[$-1060000]dd/mm/yyyy", "bbbb-mm-dd", "ge.mm.dd" })
            Check($"not guessed: {code}", new NumFmt(code).Number(45236) == null);
    }

    static void Policy()
    {
        void Refused(string path, string why) { try { CachePolicy.Validate(path); Check("refused " + why, false); } catch (CachePolicyException) { Check("refused " + why, true); } }
        Check("this laptop is approved", CachePolicy.ThisPcApproved);
        Refused(@"C:\Temp\x.xvc", "other disk"); Refused(@"D:\WORK\x.xvc", "other folder"); Refused(@"D:\ExcelViewerCache\sub\x.xvc", "sub folder"); Refused(@"S:\ExcelViewerCache\x.xvc", "network");
        Refused(@"D:\WORK\Software Development\GitHub\Excel\x.xvc", "repository"); Refused(Environment.GetEnvironmentVariable("OneDrive") + @"\x.xvc", "synced folder");
        Refused(@"D:\ExcelViewerCache\..\x.xvc", "folder escape"); Refused(@"D:\ExcelViewerCache\x.csv", "other file type"); Refused(@"D:\ExcelViewerCache\real-test.xvc", "the owner's kept copy");
        CachePolicy.Validate(CachePolicy.PathFor(@"D:\any\file.xlsx"));
        Check("cache name hides the source name", !CachePolicy.PathFor(@"D:\any\secret-name.xlsx").Contains("secret"));
        try { CacheBuilder.Build(@"D:\none.xlsx", @"C:\Temp\x.xvc", _ => { }, () => false); Check("builder refuses outside path", false); }
        catch (CachePolicyException) { Check("builder refuses outside path", true); }

        // Only CachePolicy may create, replace or delete a file in the core code.
        string dir = AppContext.BaseDirectory; while (dir != null && !Directory.Exists(Path.Combine(dir, "ExcelViewer", "Core"))) dir = Path.GetDirectoryName(dir);
        Check("found the source", dir != null);
        if (dir != null)
            foreach (var f in Directory.GetFiles(Path.Combine(dir, "ExcelViewer"), "*.cs", SearchOption.AllDirectories))
            {
                if (f.Contains(@"\obj\") || f.Contains(@"\bin\") || f.EndsWith("CachePolicy.cs")) continue;
                string s = File.ReadAllText(f);
                foreach (var bad in new[] { "FileMode.Create", "File.Create", "FileAccess.Write", "File.Move", "File.Delete", "File.WriteAll", "File.Copy", "HttpClient", "WebClient", "Socket", "SaveCopyAs", ".SaveAs(" })
                    Check($"{Path.GetFileName(f)} does not use {bad}", !s.Contains(bad));
            }
    }

    // One run by hand: the real Windows clipboard, with made-up text. Runs only if the clipboard is empty or plain text, and puts that text back.
    static int RealClipboard()
    {
        var CB = typeof(System.Windows.Forms.Clipboard); var now = System.Windows.Forms.Clipboard.GetDataObject();
        var formats = now?.GetFormats() ?? Array.Empty<string>();
        bool onlyText = formats.All(f => f is "Text" or "UnicodeText" or "System.String" or "OEMText" or "Locale");
        if (!onlyText) { Console.WriteLine("SKIPPED: the clipboard holds more than text, and could not be put back exactly"); return 3; }
        string before = System.Windows.Forms.Clipboard.ContainsText() ? System.Windows.Forms.Clipboard.GetText() : null;
        using var form = new ExcelViewer.MainForm();
        const string Want = "made-up\t$1,234.50\r\n06-Nov-2023\t12.5%";
        form.ClipboardSink(Want);
        string got = System.Windows.Forms.Clipboard.GetText();
        var after = System.Windows.Forms.Clipboard.GetDataObject().GetFormats();
        Console.WriteLine($"text came back exactly: {got == Want}");
        foreach (var f in new[] { "ExcludeClipboardContentFromMonitorProcessing", "CanIncludeInClipboardHistory", "CanUploadToCloudClipboard" }) Console.WriteLine($"{f} present: {after.Contains(f)}");
        if (before != null) System.Windows.Forms.Clipboard.SetText(before); else System.Windows.Forms.Clipboard.Clear();
        bool restored = before != null ? System.Windows.Forms.Clipboard.GetText() == before : !System.Windows.Forms.Clipboard.ContainsText();
        Console.WriteLine($"clipboard put back as it was: {restored}");
        return got == Want && restored ? 0 : 1;
    }

    static void CopyPieces()
    {
        string Txt(int r, int c) => $"r{r}c{c}";
        Check("copy one cell", CopyText.Build(new[] { (3, 2) }, Txt) == "r3c2");
        Check("copy a block: tabs and line breaks", CopyText.Build(new[] { (1, 1), (1, 2), (2, 1), (2, 2) }, Txt) == "r1c1\tr1c2\r\nr2c1\tr2c2");
        Check("copy keeps screen order, whatever the click order", CopyText.Build(new[] { (2, 2), (1, 1), (2, 1), (1, 2) }, Txt) == "r1c1\tr1c2\r\nr2c1\tr2c2");
        Check("copy keeps a gap for unselected cells", CopyText.Build(new[] { (1, 1), (2, 2) }, Txt) == "r1c1\t\r\n\tr2c2");
        Check("copy nothing", CopyText.Build(Array.Empty<(int, int)>(), Txt) == "");
        Check("a tab or line break inside a cell is quoted", CopyText.Build(new[] { (0, 0), (0, 1) }, (r, c) => c == 0 ? "a\tb" : "say \"x\"\ny") == "\"a\tb\"\t\"say \"\"x\"\"\ny\"");
    }

    // The real grid, off screen, with the clipboard swapped for a list so the owner's clipboard is not touched.
    static void CopyInWindow(string cachePath, List<(int row, int col, string text, string note)> want, CacheData c)
    {
        using var form = new ExcelViewer.MainForm { Unattended = true }; form.PlaceOffScreen();
        var got = new List<string>(); form.ClipboardSink = got.Add;
        form.Show(); form.LoadForTest(cachePath); System.Windows.Forms.Application.DoEvents();
        var g = form.Grid; int bix = c.Sheets.FindIndex(x => x.Name == "Formats");
        form.ShowSheet(bix); System.Windows.Forms.Application.DoEvents();
        Check("the grid can select single cells", g.SelectionMode == System.Windows.Forms.DataGridViewSelectionMode.RowHeaderSelect && g.MultiSelect);

        // The first rows hold dates and money. Excel's own text is the expected answer.
        var first = want.Where(w => w.row < 3).OrderBy(w => w.row).ToList(); int off = form.FirstDataRow;
        g.ClearSelection(); g[1, first[0].row - off].Selected = true;
        Check("copy returns the shown text of one cell", Clean(form.SelectionText()) == Clean(first[0].text), $"'{form.SelectionText()}' want '{first[0].text}'");
        form.HandleGridKeys(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.C);
        Check("Ctrl+C copies", got.Count == 1 && Clean(got[0]) == Clean(first[0].text));
        got.Clear(); form.CopyMenuItem.PerformClick();
        Check("right-click Copy copies", got.Count == 1 && Clean(got[0]) == Clean(first[0].text));

        g.ClearSelection(); for (int r = 0; r < 3; r++) { g[0, first[r].row - off].Selected = true; g[1, first[r].row - off].Selected = true; }
        var lines = (form.SelectionText() ?? "").Split("\r\n");
        Check("copy a block has one line per row", lines.Length == 3, $"{lines.Length} lines");
        for (int r = 0; r < Math.Min(3, lines.Length); r++)
        {
            var cells = lines[r].Split('\t');
            Check($"block row {r + 1} is tab-separated with the shown date or money", cells.Length == 2 && cells[0] == "label " + (first[r].row + 1) && Clean(cells[1]) == Clean(first[r].text), $"'{lines[r]}' want '{first[r].text}'");
        }

        // After a sort, copy must follow what is on the screen.
        int big = c.Sheets.FindIndex(x => x.Name == "Big"); form.ShowSheet(big); System.Windows.Forms.Application.DoEvents();
        form.RowMap = SortEngine.Sort(c, c.Sheets[big], 1, true, form.FirstDataRow, CancellationToken.None); form.Redraw();
        g.ClearSelection(); g[1, 0].Selected = true; string top = form.SelectionText();
        using (var rd = new SheetReader(c, c.Sheets[big])) Check("copy follows the sorted order, with its format", top == rd.Cell(form.RowMap[0], 1).Text && top.Contains('.'), $"'{top}'");
        g.ClearSelection(); got.Clear(); form.CopySelection(); Check("nothing selected copies nothing", got.Count == 0);
        Check("select-all is blocked", form.HandleGridKeys(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.A) && g.GetCellCount(System.Windows.Forms.DataGridViewElementStates.Selected) == 0);
        form.Close();
    }

    // An older or damaged copy must never stop the viewer. This is what once hung a run.
    static void BadCopies()
    {
        string path = CachePolicy.PathFor(@"D:\none\syn-bad-copy-test.xlsx"), src = Path.Combine(Path.GetTempPath(), "syn-bad-" + Guid.NewGuid().ToString("N")[..8] + ".txt");
        File.WriteAllText(src, "made-up"); var fi = new FileInfo(src);
        void Write(string magic, byte[] rest)
        {
            using var fs = CachePolicy.OpenForWrite(path); using var w = new BinaryWriter(fs);
            w.Write(System.Text.Encoding.ASCII.GetBytes(magic)); w.Write(fi.Length); w.Write(fi.LastWriteTimeUtc.Ticks); w.Write(rest);
        }
        bool Damaged() { try { CacheData.Load(path); return false; } catch (InvalidDataException) { return true; } }
        try
        {
            Write("XVCACHE2", new byte[64]); Check("an older layout is not fresh", !CacheData.HeaderMatches(path, src));
            Write(CacheData.Magic, Array.Empty<byte>()); Check("a cut-off copy is called damaged", Damaged());
            var junk = new byte[4096]; new Random(1).NextBytes(junk); Write(CacheData.Magic, junk); Check("a scrambled copy is called damaged", Damaged());
            Write(CacheData.Magic, BitConverter.GetBytes(-5L)); Check("a bad position is called damaged", Damaged());
        }
        finally { CachePolicy.Delete(path); File.Delete(src); }
    }

    // ---------- With a real hidden Excel, on a made-up workbook ----------

    static void Live()
    {
        string dir = Path.Combine(Path.GetTempPath(), "syn-viewer-test-" + Guid.NewGuid().ToString("N")[..8]); Directory.CreateDirectory(dir);
        string src = Path.Combine(dir, "syn-test.xlsx"), cachePath = CachePolicy.PathFor(src);
        var kept = new FileInfo(Path.Combine(CachePolicy.Folder, CachePolicy.KeptTestCopy)); long keptLen = kept.Exists ? kept.Length : -1; var keptTime = kept.Exists ? kept.LastWriteTimeUtc : default;
        
        try
        {
            var want = MakeWorkbook(src);
            Check("workbook made", want.Count > 20);
            long srcLen = new FileInfo(src).Length; var srcTime = new FileInfo(src).LastWriteTimeUtc;
            Check("no copy yet means not fresh", !CacheData.HeaderMatches(cachePath, src));

            var lines = new List<string>(); var sw = Stopwatch.StartNew();
            var r = CacheBuilder.Build(src, cachePath, lines.Add, () => false);
            Check("cache built", r.Code == 0, $"code {r.Code} {r.Stopped}");
            Console.WriteLine($"built in {sw.Elapsed.TotalSeconds:F1} s: formats {r.Formats}, matched {r.FormatsOk}, not carried {r.FormatsFailed}, no cells {r.FormatsUnused}");
            Check("source unchanged by the build", new FileInfo(src).Length == srcLen && new FileInfo(src).LastWriteTimeUtc == srcTime);
            Check("no half-written copy left", !File.Exists(cachePath + ".part"));
            Check("copy is in the approved folder", Path.GetDirectoryName(cachePath).Equals(CachePolicy.Folder, StringComparison.OrdinalIgnoreCase) && File.Exists(cachePath));
            var acl = new DirectoryInfo(CachePolicy.Folder).GetAccessControl();
            Check("folder is locked", acl.AreAccessRulesProtected);

            // Fast reopen, with no Excel.
            sw.Restart(); Check("fresh copy recognised", CacheData.HeaderMatches(cachePath, src));
            var c = CacheData.Load(cachePath); using var fmtView = new SheetReader(c, c.Sheets[0]); string first = fmtView.Cell(0, 0).Text; double reopen = sw.Elapsed.TotalMilliseconds;
            Check("reopen under 1 second", reopen < 1000, $"{reopen:F0} ms"); Console.WriteLine($"reopen to first cell: {reopen:F1} ms");
            foreach (var x in c.Sheets) { using var v = new SheetReader(c, x); Check($"sheet {x.Name} reads back identical", v.RecomputeHash() == x.Hash); }

            // Dates and money: every cell must show exactly what Excel shows.
            int same = 0;
            foreach (var (row, col, text, note) in want)
            {
                var shown = fmtView.Cell(row, col);
                string a = Clean(shown.Text), b = Clean(text);
                bool ok = a == b;
                Check($"shown like Excel: {note}", ok, $"viewer '{shown.Text}' Excel '{text}'");
                Check($"no warning: {note}", !shown.Warn);
                if (ok) same++;
            }
            Console.WriteLine($"dates and money: {same} of {want.Count} cells show exactly as Excel");

            foreach (var (row, text2) in new[] { (0, "0.004"), (1, "1234.5678"), (2, "123456789.5"), (3, "-0.00071"), (4, "0"), (5, "7.25") })
            { var v = fmtView.Cell(row, 2); Check("narrow General shows the full value, no warning", !v.Warn && v.Text == text2, $"'{v.Text}' warn={v.Warn}"); }
            Check("long text is not flagged", !fmtView.Cell(0, 3).Warn && fmtView.Cell(0, 3).Text.Length == 400 && fmtView.Cell(1, 3).Text.Length == 400);
            Check("General is matched to Excel", c.Formats[0].Status == FormatStatus.MatchesExcel, c.Formats[0].Status.ToString());

            Console.WriteLine($"Excel text in a hidden column: '{HiddenColText}'; in a hidden row: '{HiddenRowText}'");
            var hc = fmtView.Cell(0, 4); Check("hidden column shows its value, no warning", hc.Text == "42.5" && !hc.Warn, $"'{hc.Text}'");
            var zs = c.Sheets.Single(x => x.Name == "NoZeros"); Check("the show-zeros setting is read", zs.HideZeros);
            Check("a normal sheet still shows zeros", !c.Sheets[0].HideZeros);
            using (var zv = new SheetReader(c, zs)) for (int zr = 0; zr < 5; zr++) { var v = zv.Cell(zr, 0); Check($"zeros hidden like Excel, row {zr + 1}", Clean(v.Text) == Clean(ZeroSheet[zr]) && !v.Warn, $"viewer '{v.Text}' Excel '{ZeroSheet[zr]}'"); }

            // A format we cannot carry must say so, not guess.
            int hijriRow = LastFormatRow;
            var h = fmtView.Cell(hijriRow, 1);
            Check("unsupported format is flagged", h.Warn && h.Text.Contains('\u26A0'), $"'{h.Text}'");
            Check("sheet reports the warning", fmtView.ColumnsWithFormatWarnings() >= 1);

            // Search
            var big = c.Sheets.Single(x => x.Name == "Big"); int bigIx = c.Sheets.IndexOf(big);
            var hits = new List<Hit>(); sw.Restart(); bool all = SearchEngine.Run(c, new[] { 0, bigIx }, "needle-known-777", hits.Add, _ => { }, CancellationToken.None);
            Check("search covers everything", all); Check("search finds the known value once", hits.Count == 1 && hits[0].Sheet == bigIx && hits[0].Row == KnownRow && hits[0].Col == 3, $"{hits.Count} hits");
            Console.WriteLine($"search {big.Cells} cells: {sw.Elapsed.TotalMilliseconds:F0} ms");
            hits.Clear(); SearchEngine.Run(c, new[] { bigIx }, "NEEDLE-KNOWN", hits.Add, _ => { }, CancellationToken.None); Check("search ignores upper and lower case", hits.Count == 1);
            hits.Clear(); SearchEngine.Run(c, new[] { 0 }, "1,234,567.89", hits.Add, _ => { }, CancellationToken.None); Check("search finds a number as shown", hits.Count >= 1);
            hits.Clear(); SearchEngine.Run(c, new[] { 0 }, "1234567.89", hits.Add, _ => { }, CancellationToken.None); Check("search finds it without commas", hits.Count >= 1);
            hits.Clear(); SearchEngine.Run(c, new[] { bigIx }, "no-such-thing-zz", hits.Add, _ => { }, CancellationToken.None); Check("search with nothing to find", hits.Count == 0);
            using (var cts = new CancellationTokenSource()) { cts.Cancel(); bool threw = false; try { SearchEngine.Run(c, new[] { bigIx }, "x", _ => { }, _ => { }, cts.Token); } catch (OperationCanceledException) { threw = true; } Check("search can be stopped", threw); }

            // Sort
            var bv = new SheetReader(c, big);
            sw.Restart(); var map = SortEngine.Sort(c, big, 1, false, 1, CancellationToken.None); Console.WriteLine($"sort number column: {sw.Elapsed.TotalMilliseconds:F0} ms");
            Check("sort keeps every row but the header", map.Length == big.Rows - 1 && map.Distinct().Count() == map.Length && !map.Contains(0));
            bool asc = true, emptiesLast = true, seenEmpty = false; double prev = double.MinValue; int prevRow = -1, stable = 0;
            foreach (int row in map)
            {
                var col = bv.Column(row, 1); int i = row % big.BlockRows; byte k = col.KindAt(i);
                if (k == CellKind.Empty) { seenEmpty = true; continue; }
                if (seenEmpty) emptiesLast = false;
                if (k == CellKind.Number) { double v = col.Nums[i]; if (v < prev) asc = false; if (v == prev && row < prevRow) stable++; prev = v; prevRow = row; }
            }
            Check("sort ascending", asc); Check("sort: empty cells last", emptiesLast && seenEmpty); Check("sort: equal values keep their order", stable == 0);
            var desc = SortEngine.Sort(c, big, 1, true, 1, CancellationToken.None); double hi = double.MaxValue; bool isDesc = true;
            foreach (int row in desc) { var col = bv.Column(row, 1); int i = row % big.BlockRows; if (col.KindAt(i) == CellKind.Number) { if (col.Nums[i] > hi) isDesc = false; hi = col.Nums[i]; } else if (col.KindAt(i) == CellKind.Empty) break; }
            Check("sort descending, empties still last", isDesc && bv.Column(desc[^1], 1).KindAt(desc[^1] % big.BlockRows) == CellKind.Empty);
            var byText = SortEngine.Sort(c, big, 2, false, 1, CancellationToken.None); string pt = ""; bool tok = true;
            foreach (int row in byText.Take(5000)) { string s = bv.RawText(row, 2); if (s == null) continue; if (string.Compare(pt, s, StringComparison.OrdinalIgnoreCase) > 0) tok = false; pt = s; }
            Check("sort text A to Z", tok);
            Check("sort does not change the source", new FileInfo(src).LastWriteTimeUtc == srcTime);

            // A changed source makes the copy stale.
            File.SetLastWriteTimeUtc(src, srcTime.AddMinutes(1)); Check("changed file means the copy is not fresh", !CacheData.HeaderMatches(cachePath, src));
            File.SetLastWriteTimeUtc(src, srcTime); Check("same file again is fresh", CacheData.HeaderMatches(cachePath, src));

            CopyInWindow(cachePath, want, c);

            // Closing the viewer stops the reader, which closes its own Excel.
            fmtView.Dispose(); bv.Dispose(); CancelTest(src, cachePath);

            Check("nothing else written beside the source", Snapshot(dir).Count == 1);
        }
        finally
        {
            GC.Collect(); GC.WaitForPendingFinalizers();
            try { CachePolicy.Delete(cachePath); CachePolicy.Delete(cachePath + ".part"); } catch (IOException) { }
            try { Directory.Delete(dir, true); } catch (IOException) { }
            kept.Refresh();
            Check("the owner's kept copy is untouched", keptLen < 0 || (kept.Length == keptLen && kept.LastWriteTimeUtc == keptTime));
        }
    }

    static List<string> Snapshot(string d) => Directory.GetFiles(d).ToList();
    static string Clean(string s) => string.Concat(s.Where(c => !char.IsWhiteSpace(c) && c != '\u00A0'));

    const int KnownRow = 12345; static int LastFormatRow; static readonly List<string> ZeroSheet = new(); static string HiddenColText, HiddenRowText;

    static readonly (string fmt, double[] vals)[] Table =
    {
        ("yyyy-mm-dd", new[] { 45236.0, 1, 36526 }), ("dd/mm/yyyy", new[] { 45236.0 }), ("m/d/yyyy", new[] { 45236.0, 44927 }), ("d-mmm-yy", new[] { 45306.0, 45337, 45366, 45397, 45427, 45458, 45488, 45519, 45550, 45580, 45611, 45641 }), ("dd-mmmm-yyyy", new[] { 45550.0, 45427 }), ("ddd d mmm", new[] { 45550.0, 45551, 45552 }), ("mmmm d, yyyy", new[] { 45236.0 }), ("dddd, mmmm dd, yyyy", new[] { 45236.0 }), ("dd.mm.yyyy", new[] { 45236.0 }),
        ("h:mm AM/PM", new[] { 0.75, 0.25 }), ("hh:mm:ss", new[] { 0.5, 0.123456, 0.999999, 0.99999, 0.9999999 }), ("[h]:mm:ss", new[] { 1.5 }), ("yyyy-mm-dd hh:mm", new[] { 45236.5 }), ("m/d/yy h:mm", new[] { 45236.75 }),
        ("#,##0", new[] { 1234567.0, -1234, 0 }), ("#,##0.00", new[] { 1234567.891, -0.5, 0 }), ("$#,##0.00", new[] { 1234.5, -99.99 }), ("$#,##0.00;[Red]($#,##0.00)", new[] { 1234.5, -1234.5, 0 }), ("[$EGP] #,##0.00", new[] { 1500.25 }), ("#,##0.00\" EGP\"", new[] { 2500.0 }),
        ("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", new[] { 1234.5, -1234.5, 0 }), ("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)", new[] { 1234.0, -1234, 0 }), ("#,##0,", new[] { 1234567.0 }), ("0.0,,", new[] { 2500000.0 }),
        ("0%", new[] { 0.256, -0.1 }), ("0.00%", new[] { 0.12345, 1.5, 0 }), ("0%", new[] { -0.001, -0.004, -0.006 }), ("0.0", new[] { -0.04, -0.06 }), ("#,##0", new[] { -0.4, -0.6 }), ("$#,##0.00", new[] { -0.001 }), ("0.00E+00", new[] { -0.001 }), ("0.00;[Red]-0.00", new[] { -0.001 }), ("#,##0.00;(#,##0.00)", new[] { -0.001 }), ("0.0%", new[] { 0.5 }),
        ("0", new[] { 12.7 }), ("0.00", new[] { 1.005, 3.14159 }), ("00000", new[] { 42.0 }), ("000-00-0000", new[] { 123456789.0 }), ("0.00E+00", new[] { 12345.0, 0.00012 }), ("# ?/?", new[] { 2.5 }), ("# ??/16", new[] { 2.25 }), ("0;-0;\"zero\"", new[] { 0.0, 5, -5 }),
    };

    // Writes made-up values, lets Excel format them, and records Excel's own text as the expected answer.
    static List<(int row, int col, string text, string note)> MakeWorkbook(string path)
    {
        var want = new List<(int, int, string, string)>();
        dynamic xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        try
        {
            xl.Visible = false; xl.DisplayAlerts = false; xl.AutomationSecurity = 3;
            dynamic wb = xl.Workbooks.Add(); while ((int)wb.Worksheets.Count < 3) wb.Worksheets.Add();
            dynamic f = wb.Worksheets[1]; f.Name = "Formats";
            int row = 1;
            foreach (var (fmt, vals) in Table)
                foreach (double v in vals)
                {
                    f.Cells[row, 1].Value2 = "label " + row; f.Cells[row, 2].NumberFormat = fmt; f.Cells[row, 2].Value2 = v; row++;
                }
            int plain = row; f.Cells[row, 2].Value2 = 1234.5; row++; f.Cells[row, 2].Value2 = 0.1; row++; f.Cells[row, 2].Value2 = -7; row++;
            f.Cells[row, 2].Value2 = true; row++; f.Cells[row, 2].Formula = "=1/0"; row++; f.Cells[row, 2].Formula = "=NA()"; row++;
            f.Cells[row, 2].Value2 = "plain text"; row++; f.Cells[row, 2].NumberFormat = "@"; f.Cells[row, 2].Value2 = "00123"; row++;
            f.Cells[row, 2].NumberFormat = "0;0;0;\"[\"@\"]\""; f.Cells[row, 2].Value2 = "abc"; row++;
            f.Cells[row, 2].Formula = "=DATE(2024,2,29)"; f.Cells[row, 2].NumberFormat = "yyyy-mm-dd"; row++;
            f.Columns[2].ColumnWidth = 40;
            // Narrow General numbers and long text: Excel shortens these on screen, which must not count as a format failure.
            f.Cells[1, 3].Value2 = 0.004; f.Cells[2, 3].Value2 = 1234.5678; f.Cells[3, 3].Value2 = 123456789.5; f.Cells[4, 3].Value2 = -0.00071; f.Cells[5, 3].Value2 = 0; f.Cells[6, 3].Value2 = 7.25; f.Columns[3].ColumnWidth = 2;
            // Numbers in a hidden column and a hidden row.
            f.Cells[1, 5].Value2 = 42.5; f.Cells[2, 5].Value2 = 0; f.Columns[5].Hidden = true;
            f.Cells[60, 6].Value2 = 77; f.Rows[60].Hidden = true; f.Columns[6].ColumnWidth = 10;
            HiddenColText = (string)f.Cells[1, 5].Text; HiddenRowText = (string)f.Cells[60, 6].Text;
            f.Cells[1, 4].NumberFormat = "@"; f.Cells[1, 4].Value2 = new string('x', 400); f.Cells[2, 4].Value2 = new string('y', 400); f.Columns[4].ColumnWidth = 3;
            int checkEnd = row;
            // A Hijri date format, which the viewer does not carry.
            f.Cells[row, 2].NumberFormat = "[$-1060000]dd/mm/yyyy"; f.Cells[row, 2].Value2 = 45236; LastFormatRow = row - 1;
            for (int r = 1; r < checkEnd; r++) { string t = (string)f.Cells[r, 2].Text; want.Add((r - 1, 1, t, "row " + r + " format " + (string)f.Cells[r, 2].NumberFormat)); }

            // A bigger made-up sheet: numbers with gaps, text, one known value.
            dynamic b = wb.Worksheets[2]; b.Name = "Big"; const int N = 50000;
            var arr = new object[N + 1, 4]; arr[0, 0] = "Id"; arr[0, 1] = "Amount"; arr[0, 2] = "Group"; arr[0, 3] = "Note"; var rng = new Random(11);
            for (int i = 1; i <= N; i++) { arr[i, 0] = (double)i; arr[i, 1] = i % 50 == 0 ? null : (object)(double)rng.Next(0, 500); arr[i, 2] = "group-" + rng.Next(100); arr[i, 3] = "note " + i; }
            arr[KnownRow, 3] = "needle-known-777";
            b.Range["A1:D" + (N + 1)].Value2 = arr; b.Range["B2:B" + (N + 1)].NumberFormat = "#,##0.00";
            // A sheet with Excel's "show zeros" setting off.
            dynamic z = wb.Worksheets[3]; z.Name = "NoZeros"; z.Cells[1, 1].Value2 = 0; z.Cells[2, 1].Value2 = 5; z.Cells[3, 1].Value2 = 0; z.Cells[3, 1].NumberFormat = "#,##0.00"; z.Cells[4, 1].Value2 = 2.5; z.Cells[4, 1].NumberFormat = "#,##0.00"; z.Cells[5, 1].Value2 = 0.001; z.Cells[5, 1].NumberFormat = "0"; z.Columns[1].ColumnWidth = 20;
            z.Activate(); xl.ActiveWindow.DisplayZeros = false;
            for (int r = 1; r <= 5; r++) ZeroSheet.Add((string)z.Cells[r, 1].Text);
            f.Activate();
            wb.SaveAs(path, 51); wb.Close(false);
            Marshal.FinalReleaseComObject(z); Marshal.FinalReleaseComObject(b); Marshal.FinalReleaseComObject(f); Marshal.FinalReleaseComObject(wb);
        }
        finally { xl.Quit(); Marshal.FinalReleaseComObject(xl); xl = null; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); }
        return want;
    }

    static void CancelTest(string src, string cachePath)
    {
        string exe = Path.Combine(AppContext.BaseDirectory, "ExcelViewer.exe");
        Check("viewer program found", File.Exists(exe)); if (!File.Exists(exe)) return;
        CachePolicy.Delete(cachePath);
        var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        var psi = new ProcessStartInfo(exe) { RedirectStandardOutput = true, RedirectStandardInput = true, UseShellExecute = false, CreateNoWindow = true }; psi.ArgumentList.Add("--build"); psi.ArgumentList.Add(src);
        using var p = Process.Start(psi); string line;
        while ((line = p.StandardOutput.ReadLine()) != null && line != "OPENED") { }
        Check("reader reports opening", line == "OPENED");
        p.StandardInput.Close();
        p.StandardOutput.ReadToEnd();
        bool ended = p.WaitForExit(60000);
        Check("closing the viewer stops the reader", ended && p.ExitCode == 2, ended ? $"code {p.ExitCode}" : "still running");
        Check("cancel keeps no copy", !File.Exists(cachePath) && !File.Exists(cachePath + ".part"));
        var sw = Stopwatch.StartNew(); bool gone = false;
        while (sw.Elapsed.TotalSeconds < 30) { if (!Process.GetProcessesByName("EXCEL").Any(x => !before.Contains(x.Id))) { gone = true; break; } Thread.Sleep(250); }
        Check("the reader's Excel left by itself", gone);
    }
}

