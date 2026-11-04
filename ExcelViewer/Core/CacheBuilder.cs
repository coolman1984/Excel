// Builds the fast copy from what Excel reads through COM. Excel is the only reader of the source.
// Runs in its own small process. Reports only times, sizes and counts: never a cell, a name or a message
// from Excel, since an error message could quote file contents.
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelViewer.Core;

public sealed class BuildResult
{
    public int Code; public string Stopped; public double ExcelStartSec, OpenSec, FormatSec, ReadSec, VerifySec; public int Sheets; public long Cells;
    public int Formats, FormatsOk, FormatsFailed, FormatsUnused; public long UnreadRuns; public long CacheBytes; public long PeakExcelMB; public double PeakCommit;
}

public static class Machine
{
    [StructLayout(LayoutKind.Sequential)] struct MS { public uint len, load; public ulong tp, ap, tf, af, tv, av, ae; }
    [DllImport("kernel32.dll")] static extern bool GlobalMemoryStatusEx(ref MS m);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    public const double StopAt = 96; // the owner's rule: only stop here (2026-11-04)
    public static double Commit() { var m = new MS { len = (uint)Marshal.SizeOf<MS>() }; GlobalMemoryStatusEx(ref m); return 100.0 * (m.tf - m.af) / m.tf; }
    public static long PrivateMB(int id) { if (id == 0) return 0; try { using var p = Process.GetProcessById(id); return p.PrivateMemorySize64 / 1048576; } catch (ArgumentException) { return 0; } }
}

public static class CacheBuilder
{
    const long CallBudget = 300_000; const double FormatSecondsBudget = 180;

    sealed class Sample { public int Sheet, Row, Col; public double Num; public string Text; }
    sealed class SampleSet { public List<Sample> List = new(); public bool Neg, Zero, Frac, Pos; public int TextCount; }

    public static BuildResult Build(string source, string cachePath, Action<string> report, Func<bool> cancelled)
    {
        var res = new BuildResult();
        CachePolicy.Validate(cachePath); // refuse before Excel even starts
        string part = cachePath + ".part"; CachePolicy.Validate(part); CachePolicy.EnsureFolderLocked();
        if (Machine.Commit() >= Machine.StopAt) { res.Code = 90; res.Stopped = "memory"; return res; }

        var src = new FileInfo(source); long srcLen = src.Length, srcTicks = src.LastWriteTimeUtc.Ticks;
        int xpid = 0; bool sampling = true, stop = false;
        var sampler = new Thread(() =>
        {
            while (Volatile.Read(ref sampling))
            {
                res.PeakExcelMB = Math.Max(res.PeakExcelMB, Machine.PrivateMB(xpid)); double c = Machine.Commit(); res.PeakCommit = Math.Max(res.PeakCommit, c);
                if (c >= Machine.StopAt) stop = true;
                Thread.Sleep(500);
            }
        }) { IsBackground = true }; sampler.Start();

        var sw = Stopwatch.StartNew();
        report("OPENING");
        dynamic xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        dynamic wb = null;
        try
        {
            xl.Visible = false; xl.DisplayAlerts = false; xl.AutomationSecurity = 3; xl.ScreenUpdating = false; xl.EnableEvents = false;
            Machine.GetWindowThreadProcessId(new IntPtr((long)xl.Hwnd), out uint pid); xpid = (int)pid;
            res.ExcelStartSec = sw.Elapsed.TotalSeconds; sw.Restart();
            wb = xl.Workbooks.Open(source, 0, true);
            res.OpenSec = sw.Elapsed.TotalSeconds; report("OPENED");
            if (stop || cancelled()) { res.Code = stop ? 90 : 2; res.Stopped = stop ? "memory" : "cancelled"; return res; }

            var cache = new CacheData { SourceLength = srcLen, SourceTicks = srcTicks, Date1904 = (bool)wb.Date1904 };
            cache.Formats.Add(new FormatEntry { Code = "General" });
            var ids = new Dictionary<string, int>(StringComparer.Ordinal) { ["General"] = 0 };

            // Size of each sheet, with empty formatted space trimmed.
            var sheets = new List<dynamic>();
            foreach (dynamic ws in wb.Worksheets)
            {
                dynamic u = ws.UsedRange; var d = new SheetInfo { Name = (string)ws.Name, Hidden = (int)ws.Visible != -1, FirstRow = u.Row, FirstCol = u.Column, Rows = u.Rows.Count, Cols = u.Columns.Count, Hash = 14695981039346656037UL };
                try
                {
                    dynamic lastR = ws.Cells.Find("*", Type.Missing, -4123, Type.Missing, 1, 2);
                    if (lastR == null) { d.Rows = 0; d.Cols = 0; }
                    else
                    {
                        dynamic lastC = ws.Cells.Find("*", Type.Missing, -4123, Type.Missing, 2, 2);
                        d.Rows = Math.Max(0, (int)lastR.Row - d.FirstRow + 1); d.Cols = Math.Max(0, (int)lastC.Column - d.FirstCol + 1);
                    }
                }
                catch (COMException) { }
                d.BlockRows = d.Cols == 0 ? 1 : Math.Clamp(20000 / d.Cols, 50, 2000);
                d.HideZeros = ZerosHidden(wb, ws);
                cache.Sheets.Add(d); sheets.Add(ws); res.Cells += d.Cells;
            }
            res.Sheets = cache.Sheets.Count;
            long free = new DriveInfo(Path.GetPathRoot(CachePolicy.Folder)).AvailableFreeSpace;
            if (res.Cells * 12 > free - (1L << 30)) { res.Code = 5; res.Stopped = "disk"; return res; }

            // Number formats, as runs of rows per column. A column with one format costs one call.
            sw.Restart(); long calls = 0; report("FORMATS");
            for (int s = 0; s < sheets.Count; s++)
            {
                var d = cache.Sheets[s]; dynamic ws = sheets[s]; d.ColWidths = new float[d.Cols]; d.RunStart = new int[d.Cols][]; d.RunFmt = new int[d.Cols][];
                for (int k = 0; k < d.Cols; k++)
                {
                    if (stop || cancelled()) { res.Code = stop ? 90 : 2; res.Stopped = stop ? "memory" : "cancelled"; return res; }
                    string col = Letters(d.FirstCol + k);
                    try { d.ColWidths[k] = (float)(double)ws.Columns[d.FirstCol + k].ColumnWidth; } catch (COMException) { d.ColWidths[k] = 10; } catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { d.ColWidths[k] = 10; }
                    var runs = new List<(int start, int fmt)>();
                    ReadRuns((object)ws, col, d.FirstRow, 0, d.Rows - 1, runs, cache, ids, ref calls, sw);
                    d.RunStart[k] = runs.Select(r => r.start).ToArray(); d.RunFmt[k] = runs.Select(r => r.fmt).ToArray();
                }
            }
            res.FormatSec = sw.Elapsed.TotalSeconds;

            // The values.
            sw.Restart(); long done = 0; int lastPct = -1;
            var sets = new Dictionary<int, SampleSet>();
            using (var fs = CachePolicy.OpenForWrite(part))
            using (var w = new BinaryWriter(fs, Encoding.UTF8))
            {
                w.Write(Encoding.ASCII.GetBytes(CacheData.Magic)); w.Write(srcLen); w.Write(srcTicks);
                for (int s = 0; s < sheets.Count && !stop; s++)
                {
                    var d = cache.Sheets[s]; dynamic ws = sheets[s]; string c1 = Letters(d.FirstCol), c2 = Letters(d.FirstCol + d.Cols - 1);
                    for (int r = 0; r < d.Rows; r += d.BlockRows)
                    {
                        if (stop || cancelled()) { stop = true; break; }
                        int n = Math.Min(d.BlockRows, d.Rows - r);
                        object[,] a = ReadBlock((object)ws, $"{c1}{d.FirstRow + r}:{c2}{d.FirstRow + r + n - 1}");
                        var off = new long[d.Cols + 1];
                        for (int k = 0; k < d.Cols; k++)
                        {
                            var kinds = new byte[n]; double[] nums = null; string[] text = null;
                            for (int i = 0; i < n; i++)
                                switch (a[i + 1, k + 1])
                                {
                                    case null: break;
                                    case double x: kinds[i] = CellKind.Number; (nums ??= new double[n])[i] = x; break;
                                    case string t: kinds[i] = CellKind.Text; (text ??= new string[n])[i] = t; break;
                                    case bool b: kinds[i] = CellKind.Bool; (nums ??= new double[n])[i] = b ? 1 : 0; break;
                                    case int e: kinds[i] = CellKind.Error; (nums ??= new double[n])[i] = e; break;
                                    default: kinds[i] = CellKind.Number; (nums ??= new double[n])[i] = Convert.ToDouble(a[i + 1, k + 1], CultureInfo.InvariantCulture); break;
                                }
                            Collect(sets, d, s, k, r, n, kinds, nums, text, cache);
                            var cc = ColumnChunk.Build(n, kinds, nums, text);
                            d.Hash = cc.Mix(d.Hash, n);
                            w.Flush(); off[k] = fs.Position; cc.Write(w);
                        }
                        w.Flush(); off[d.Cols] = fs.Position; d.Chunks.Add(off); done += (long)n * d.Cols;
                        int pct = res.Cells == 0 ? 100 : (int)(done * 100 / res.Cells);
                        if (pct != lastPct) { lastPct = pct; report("READ " + pct); }
                    }
                }
                if (!stop)
                {
                    res.ReadSec = sw.Elapsed.TotalSeconds; sw.Restart(); report("CHECKING");
                    cache.Names = ReadNames(xl);
                    Verify(xl, sheets, cache, sets, report);
                    res.VerifySec = sw.Elapsed.TotalSeconds;
                    w.Flush(); long at = fs.Position; cache.WriteDirectory(w); w.Write(at);
                }
            }
            if (stop) { CachePolicy.Delete(part); res.Code = cancelled() ? 2 : 90; res.Stopped = cancelled() ? "cancelled" : "memory"; return res; }
            CachePolicy.ReplaceWith(part, cachePath);
            res.CacheBytes = new FileInfo(cachePath).Length;
            res.Formats = cache.Formats.Count;
            res.FormatsOk = cache.Formats.Count(f => f.Status == FormatStatus.MatchesExcel);
            res.FormatsFailed = cache.Formats.Count(f => f.Status == FormatStatus.CannotCarry);
            res.FormatsUnused = cache.Formats.Count(f => f.Status == FormatStatus.NotChecked);
            res.UnreadRuns = cache.Sheets.Sum(d => (long)d.RunFmt.Sum(a => a.Count(x => x < 0)));
            return res;
        }
        catch (Exception e) when (e is COMException or InvalidCastException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or IOException)
        {
            res.Code = 1; res.Stopped = e.GetType().Name; // the type only
            try { CachePolicy.Delete(part); } catch (IOException) { }
            return res;
        }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch (COMException) { }
            wb = null;
            try { xl.Quit(); } catch (COMException) { }
            Marshal.FinalReleaseComObject(xl); xl = null;
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            Volatile.Write(ref sampling, false); // Excel is closed, never killed
        }
    }

    // The sheet's "show zeros" setting. If it cannot be read, zeros that Excel hides will fail the check and be marked.
    static bool ZerosHidden(dynamic wb, dynamic ws)
    {
        try { return !(bool)wb.Windows[1].SheetViews[(int)ws.Index].DisplayZeros; }
        catch (Exception e) when (e is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or InvalidCastException) { return false; }
    }

    public static string Letters(int n) { var s = ""; while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; } return s; }

    static object[,] ReadBlock(object sheet, string addr)
    {
        dynamic ws = sheet;
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                object v = ws.Range[addr].Value2;
                if (v is object[,] a) return a;
                var one = (object[,])Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 }); one[1, 1] = v; return one;
            }
            catch (COMException) when (attempt < 3) { Thread.Sleep(500); } // Excel busy for a moment
        }
    }

    static void ReadRuns(object sheet, string col, int firstRow, int a, int b, List<(int start, int fmt)> runs, CacheData cache, Dictionary<string, int> ids, ref long calls, Stopwatch sw)
    {
        dynamic ws = sheet;
        if (calls >= CallBudget || sw.Elapsed.TotalSeconds > FormatSecondsBudget) { Add(runs, a, -1); return; }
        calls++;
        object f = ws.Range[$"{col}{firstRow + a}:{col}{firstRow + b}"].NumberFormat;
        if (f is string s)
        {
            if (!ids.TryGetValue(s, out int id)) { id = cache.Formats.Count; ids[s] = id; cache.Formats.Add(new FormatEntry { Code = s }); }
            Add(runs, a, id); return;
        }
        if (a == b) { Add(runs, a, -1); return; }
        int mid = (a + b) / 2;
        ReadRuns(sheet, col, firstRow, a, mid, runs, cache, ids, ref calls, sw);
        ReadRuns(sheet, col, firstRow, mid + 1, b, runs, cache, ids, ref calls, sw);
    }

    static void Add(List<(int start, int fmt)> runs, int start, int fmt) { if (runs.Count == 0 || runs[^1].fmt != fmt) runs.Add((start, fmt)); }

    // Keeps a few cells of each format, covering minus, zero and fractions, for the check against Excel.
    static void Collect(Dictionary<int, SampleSet> sets, SheetInfo d, int s, int k, int r0, int n, byte[] kinds, double[] nums, string[] text, CacheData cache)
    {
        var st = d.RunStart[k]; var fm = d.RunFmt[k]; int ri = 0;
        for (int i = 0; i < n; i++)
        {
            byte kd = kinds[i]; if (kd != CellKind.Number && kd != CellKind.Text) continue;
            int row = r0 + i;
            while (ri + 1 < st.Length && st[ri + 1] <= row) ri++;
            while (ri > 0 && st[ri] > row) ri--;
            int id = fm[ri]; if (id < 0) continue;
            if (!sets.TryGetValue(id, out var set)) sets[id] = set = new SampleSet();
            if (kd == CellKind.Number)
            {
                double v = nums[i]; bool neg = v < 0, zero = v == 0, frac = v != Math.Floor(v), pos = v > 0;
                // A long run of one kind, such as zeros, must not use up the samples.
                bool newKind = (neg && !set.Neg) || (zero && !set.Zero) || (frac && !set.Frac) || (pos && !set.Pos);
                int numbers = set.List.Count(x => x.Text == null);
                if ((newKind && numbers < 16) || (numbers < 6 && !zero))
                {
                    set.List.Add(new Sample { Sheet = s, Row = d.FirstRow + row, Col = d.FirstCol + k, Num = v });
                    set.Neg |= neg; set.Zero |= zero; set.Frac |= frac; set.Pos |= pos;
                }
            }
            else if (set.TextCount < 2)
            {
                set.TextCount++; set.List.Add(new Sample { Sheet = s, Row = d.FirstRow + row, Col = d.FirstCol + k, Text = text[i] });
            }
        }
    }

    // A format counts only if our display equals Excel's own text on real cells of that format.
    static void Verify(dynamic xl, List<dynamic> sheets, CacheData cache, Dictionary<int, SampleSet> sets, Action<string> report)
    {
        foreach (var (id, set) in sets)
        {
            var e = cache.Formats[id]; if (set.List.Count == 0) continue;
            var first = set.List[0]; string local = null;
            try { local = (string)sheets[first.Sheet].Cells[first.Row, first.Col].NumberFormatLocal; } catch (COMException) { }
            var excel = new List<(Sample s, string text)>();
            foreach (var s in set.List)
            {
                string t = null;
                try { t = (string)sheets[s.Sheet].Cells[s.Row, s.Col].Text; } catch (COMException) { } catch (InvalidCastException) { }
                bool blankBecauseHidden = t != null && t.Length == 0 && s.Text == null && IsHidden((object)sheets[s.Sheet], s) && !(s.Num == 0 && cache.Sheets[s.Sheet].HideZeros);
                if (t != null && (blankBecauseHidden || (t.Length > 0 && t.All(c => c == '#'))) && s.Text == null && local != null)
                {
                    // The column is too narrow. Ask Excel to format the value without the cell.
                    try { t = (string)xl.WorksheetFunction.Text(s.Num, local); } catch (COMException) { t = null; } catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { t = null; }
                }
                if (t != null) excel.Add((s, t));
            }
            e.Samples = excel.Count;
            if (excel.Count == 0) { e.Status = FormatStatus.CannotCarry; report($"NOTCARRIED shape={Shape(e.Code)} noSamplesReadable"); continue; }
            bool ok = false;
            foreach (var code in new[] { e.Code, local }.Where(x => !string.IsNullOrEmpty(x)).Distinct())
            {
                var f = new NumFmt(code, cache.Date1904, cache.Names);
                if (excel.All(p => Same(f, code, p.s, p.text, cache.Sheets[p.s.Sheet].HideZeros)) && NamesAgree(xl, f, local ?? code, cache.Date1904)) { e.Code = code; ok = true; break; }
            }
            e.Status = ok ? FormatStatus.MatchesExcel : FormatStatus.CannotCarry;
            if (!ok) report($"NOTCARRIED shape={Shape(e.Code)} localSame={string.Equals(e.Code, local, StringComparison.Ordinal)} numberSamples={excel.Count(p => p.s.Text == null)} textSamples={excel.Count(p => p.s.Text != null)}");
            if (!ok) { var ff = new NumFmt(e.Code, cache.Date1904, cache.Names); foreach (var p in excel.Where(p => !Same(ff, e.Code, p.s, p.text, cache.Sheets[p.s.Sheet].HideZeros))) report($"MISMATCH sheet#{p.s.Sheet + 1} zero={p.s.Num == 0} hideZerosRead={cache.Sheets[p.s.Sheet].HideZeros} {Where((object)sheets[p.s.Sheet], p.s)} kind={(p.s.Text == null ? "number" : "text")} ours={Mask(p.s.Text == null ? ff.Number(p.s.Num)?.Text : ff.Text(p.s.Text))} excel={Mask(p.text)}"); }
        }
    }

    // Names as Excel writes them, for all 12 months and 7 days.
    static DateNames ReadNames(dynamic xl)
    {
        try
        {
            string[] Ask(string fmt, double first, double step, int n) { var a = new string[n]; for (int i = 0; i < n; i++) a[i] = (string)xl.WorksheetFunction.Text(first + step * i, fmt); return a; }
            // Serials: the 15th of each month of 2024, and a run of days from Sunday 2024-01-07.
            var mid = Enumerable.Range(1, 12).Select(m => (new DateTime(2024, m, 15) - new DateTime(1899, 12, 30)).TotalDays).ToArray();
            string[] Months(string fmt) => mid.Select(v => (string)xl.WorksheetFunction.Text(v, fmt)).ToArray();
            double sunday = (new DateTime(2024, 1, 7) - new DateTime(1899, 12, 30)).TotalDays;
            return new DateNames { MonthAbbr = Months("mmm"), MonthFull = Months("mmmm"), DayAbbr = Ask("ddd", sunday, 1, 7), DayFull = Ask("dddd", sunday, 1, 7) };
        }
        catch (Exception e) when (e is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or InvalidCastException) { return null; }
    }

    // A format with month or day names must agree with Excel across the whole year, not only on sampled cells.
    static bool NamesAgree(dynamic xl, NumFmt f, string code, bool d1904)
    {
        if (!f.UsesNames) return true;
        for (int i = 0; i < 19; i++)
        {
            var day = i < 12 ? new DateTime(2024, i + 1, 15) : new DateTime(2024, 1, 7 + i - 12);
            double v = (day - new DateTime(d1904 ? 1904 : 1899, d1904 ? 1 : 12, d1904 ? 1 : 30)).TotalDays;
            string t;
            try { t = (string)xl.WorksheetFunction.Text(v, code); }
            catch (Exception e) when (e is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or InvalidCastException) { return false; }
            var s = f.Number(v);
            if (s == null || Norm(s.Value.Text) != Norm(t)) return false;
        }
        return true;
    }

    static bool Same(NumFmt f, string code, Sample s, string excelText, bool hideZeros)
    {
        if (s.Text == null && s.Num == 0 && hideZeros) return Norm(excelText).Length == 0;
        if (s.Text != null)
        {
            string ours = Norm(f.Text(s.Text)), theirs = Norm(excelText);
            if (ours == theirs) return true;
            if (theirs.Length > 0 && theirs.All(c => c == '#')) return true; // long text in a narrow cell: Excel shows only # here
            return theirs.Length >= 200 && ours != null && ours.StartsWith(theirs, StringComparison.Ordinal); // Excel cuts off very long text
        }
        var shown = f.Number(s.Num); if (shown == null) return false;
        string a = Norm(shown.Value.Text), b = Norm(excelText);
        if (a == b) return true;
        if (a.All(c => c == '#') && b.All(c => c == '#')) return true;
        return (code.Equals("General", StringComparison.OrdinalIgnoreCase) || code.Length == 0) && GeneralAgrees(excelText, s.Num);
    }

    // In General, Excel rounds to fit the column. Excel's text must be a correct rounding of the same value.
    public static bool GeneralAgrees(string text, double v)
    {
        text = text.Trim(); var nf = CultureInfo.CurrentCulture.NumberFormat;
        if (!double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double x)) return false;
        int e = text.IndexOfAny(new[] { 'E', 'e' }); string mant = e < 0 ? text : text[..e];
        int dot = mant.IndexOf(nf.NumberDecimalSeparator, StringComparison.Ordinal); int decimals = dot < 0 ? 0 : mant.Length - dot - nf.NumberDecimalSeparator.Length;
        int exp = e < 0 ? 0 : int.Parse(text[(e + 1)..], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        return Math.Abs(x - v) <= 0.5005 * Math.Pow(10, exp - decimals) + 1e-12 * Math.Abs(v);
    }

    static bool IsHidden(object sheet, Sample s)
    {
        dynamic ws = sheet;
        try { return (bool)ws.Columns[s.Col].Hidden || (bool)ws.Rows[s.Row].Hidden; }
        catch (Exception e) when (e is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or InvalidCastException) { return false; }
    }

    static string Where(object sheet, Sample s)
    {
        dynamic ws = sheet;
        try { return $"colHidden={(bool)ws.Columns[s.Col].Hidden} rowHidden={(bool)ws.Rows[s.Row].Hidden} colWidth={(double)ws.Columns[s.Col].ColumnWidth:F1}"; }
        catch (Exception e) when (e is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or InvalidCastException) { return "where=unknown"; }
    }

    // The shape of a text with every digit and letter hidden, so a report can never quote a value.
    static string Mask(string v)
    {
        if (v == null) return "(none)";
        var sb = new StringBuilder(); foreach (char c in v.Length > 30 ? v[..30] : v) sb.Append(char.IsDigit(c) ? '9' : char.IsLetter(c) ? 'A' : c == ' ' ? '_' : char.IsControl(c) ? '^' : c);
        return sb + "/len" + v.Length;
    }

    // The shape of a format with any words removed, so a report can never quote file contents.
    static string Shape(string code)
    {
        var sb = new StringBuilder(); bool q = false;
        foreach (char c in code ?? "")
        {
            if (c == '"') { if (!q) sb.Append("<words>"); q = !q; continue; }
            if (q) continue;
            sb.Append(char.IsLetter(c) && !"yYmMdDhHsSeEaApPGgnNrRlL".Contains(c) ? '?' : c);
        }
        return sb.ToString();
    }

    static string Norm(string s) => s == null ? null : string.Concat(s.Where(c => !char.IsWhiteSpace(c) && c != ' '));
}
