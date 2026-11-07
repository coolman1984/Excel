using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ExcelViewer.Core;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("ExcelApp.Tests")]

namespace ExcelApp;

public sealed record StyleRun(int From, int To, int? Fill, int? FontColor, bool? Bold, int? Align);
public sealed class SheetStyles
{
    public Dictionary<int, List<StyleRun>> Columns = new();
    public List<(int From, int To)> HiddenRows = new(), HiddenCols = new();
    public int FrozenRows, FrozenCols;
}
public sealed class StyleSidecar { public Dictionary<int, SheetStyles> Sheets = new(); public bool Partial; }

public static class StyleStore
{
    public const string Magic = "XVSTYLE1";
    public const int CallBudget = 200000;
    public const double SecondsBudget = 120;
    // The synthetic 4-million-cell gate reached the 120-second budget with zero COM read errors.
    // Over 90 seconds: off by default. The host may opt in per build/save child.
    public static bool Enabled { get; set; } = false;
    public static int MaxCalls { get; set; } = CallBudget;
    public static double MaxSeconds { get; set; } = SecondsBudget;
    public static int LastErrors { get; private set; }
    public static long LastCalls { get; private set; }
    public static double LastSeconds { get; private set; }

    sealed class Budget
    {
        public readonly Stopwatch Clock = Stopwatch.StartNew(); public int Calls; public bool Partial, Exhausted; public int Errors;
        readonly int callsLimit = Math.Clamp(MaxCalls, 0, CallBudget);
        readonly double secondsLimit = double.IsFinite(MaxSeconds) ? Math.Clamp(MaxSeconds, 0, SecondsBudget) : SecondsBudget;
        public bool Take() {
            if (Machine.Commit() >= Machine.StopAt) throw new OperationCanceledException();
            if (Calls >= callsLimit || Clock.Elapsed.TotalSeconds >= secondsLimit) { Partial = true; Exhausted = true; return false; }
            Calls++; return true;
        }
    }
    public static void AfterRead(object excelApp, object workbook, CacheData cache)
    {
        if (!Enabled) return;
        dynamic wb = workbook;
        string source = (string)wb.FullName;
        var styles = Read(excelApp, workbook, cache);
        Write(source, cache, styles);
    }

    // Also used by offline sidecar round-trip tests; workbook files never pass through this writer.
    internal static void Write(string source, CacheData cache, StyleSidecar styles)
    {
        string path = CachePolicy.SidecarPathFor(source), part = path + ".part";
        if (Machine.Commit() >= Machine.StopAt) throw new OperationCanceledException();
        try {
            using (var stream = CachePolicy.OpenForWrite(part))
            using (var w = new BinaryWriter(stream, Encoding.UTF8)) {
                w.Write(Encoding.ASCII.GetBytes(Magic)); w.Write(cache.SourceLength); w.Write(cache.SourceTicks); w.Write(styles.Partial);
                w.Write(styles.Sheets.Count);
                foreach (var (index, sheet) in styles.Sheets) {
                    if (Machine.Commit() >= Machine.StopAt) throw new OperationCanceledException();
                    w.Write(index); w.Write(sheet.FrozenRows); w.Write(sheet.FrozenCols);
                    WriteRanges(w, sheet.HiddenRows); WriteRanges(w, sheet.HiddenCols);
                    w.Write(sheet.Columns.Count);
                    foreach (var (col, runs) in sheet.Columns) {
                        w.Write(col); w.Write(runs.Count);
                        foreach (var r in runs) { w.Write(r.From); w.Write(r.To); WriteInt(w, r.Fill); WriteInt(w, r.FontColor); WriteInt(w, r.Bold.HasValue ? (r.Bold.Value ? 1 : 0) : null); WriteInt(w, r.Align); }
                    }
                }
            }
            CachePolicy.ReplaceWith(part, path);
        }
        finally { CachePolicy.Delete(part); }
    }
    static void WriteInt(BinaryWriter w, int? n) { w.Write(n.HasValue); if (n.HasValue) w.Write(n.Value); }
    static int? ReadInt(BinaryReader r) => r.ReadBoolean() ? r.ReadInt32() : null;
    static void WriteRanges(BinaryWriter w, List<(int From, int To)> ranges) { w.Write(ranges.Count); foreach (var r in ranges) { w.Write(r.From); w.Write(r.To); } }

    /// <summary>Reads only properties of the supplied workbook, with one shared workbook budget.</summary>
    public static StyleSidecar Read(object excelApp, object workbook, CacheData cache)
    {
        dynamic wb = workbook;
        var result = new StyleSidecar(); var budget = new Budget();
        int active = 0;
        try {
            try { active = (int)wb.ActiveSheet.Index; } catch (COMException) { }
            for (int s = 0; s < cache.Sheets.Count; s++) {
                var d = cache.Sheets[s]; var style = new SheetStyles(); result.Sheets[s] = style;
                dynamic ws = wb.Worksheets[s + 1];
                try {
                    // Hidden sheets cannot be activated. Leave their pane information unknown.
                    if (!d.Hidden && budget.Take()) {
                        ws.Activate(); dynamic window = wb.Windows[1];
                        try { if ((bool)window.FreezePanes) { style.FrozenRows = (int)window.SplitRow; style.FrozenCols = (int)window.SplitColumn; } }
                        finally { Release(window); }
                    }
                    var hidden = Probe(0, d.Rows - 1, (a, b) => {
                        dynamic r = ws.Rows[$"{d.FirstRow + a}:{d.FirstRow + b}"];
                        try { return Hidden((object)r, true, budget); } finally { Release(r); }
                    }, budget);
                    style.HiddenRows = hidden.Where(r => r.Value == 1).Select(r => (r.From, r.To)).ToList();
                    hidden = Probe(0, d.Cols - 1, (a, b) => {
                        dynamic r = ws.Columns[$"{CacheBuilder.Letters(d.FirstCol + a)}:{CacheBuilder.Letters(d.FirstCol + b)}"];
                        try { return Hidden((object)r, false, budget); } finally { Release(r); }
                    }, budget);
                    style.HiddenCols = hidden.Where(r => r.Value == 1).Select(r => (r.From, r.To)).ToList();
                    for (int c = 0; c < d.Cols; c++) {
                        if (budget.Exhausted) break;
                        string col = CacheBuilder.Letters(d.FirstCol + c);
                        var props = new List<PropRun>[4];
                        for (int p = 0; p < 4; p++) {
                            int property = p;
                            props[p] = Probe(0, d.Rows - 1, (a, b) => {
                                dynamic range = ws.Range[$"{col}{d.FirstRow + a}:{col}{d.FirstRow + b}"];
                                try {
                                    // Excel's Color can return zero for mixed colours; ColorIndex is the mixed-value sentinel.
                                    if (property == 0) { object index = range.Interior.ColorIndex; if (Mixed(index)) return null; if (Convert.ToInt32(index) == -4142) return -1; if (!budget.Take()) return null; return (object)range.Interior.Color; }
                                    if (property == 1) { object index = range.Font.ColorIndex; if (Mixed(index)) return null; if (Convert.ToInt32(index) == -4105) return -1; if (!budget.Take()) return null; return (object)range.Font.Color; }
                                    return property == 2 ? (object)range.Font.Bold : (object)range.HorizontalAlignment;
                                }
                                finally { Release(range); }
                            }, budget);
                        }
                        var edges = props.SelectMany(r => r.SelectMany(v => new[] { v.From, v.To + 1 })).Append(0).Append(d.Rows).Distinct().OrderBy(x => x).ToArray();
                        var runs = new List<StyleRun>(); var indexes = new int[4];
                        for (int i = 0; i + 1 < edges.Length; i++) {
                            int row = edges[i];
                            int? At(int p) { while (indexes[p] < props[p].Count && props[p][indexes[p]].To < row) indexes[p]++; return indexes[p] < props[p].Count && props[p][indexes[p]].From <= row ? props[p][indexes[p]].Value : null; }
                            int? fill = At(0), font = At(1), bold = At(2), align = At(3);
                            fill = fill >= 0 ? SaveBack.Bgr(fill.Value) : null; font = font > 0 ? SaveBack.Bgr(font.Value) : null;
                            int? alignment = align switch { -4131 => 0, -4152 => 1, -4108 => 2, _ => null };
                            // Plain defaults remain sparse; gaps use the grid defaults.
                            bool? isBold = bold == 1 ? true : null;
                            if (fill == null && font == null && isBold == null && alignment == null) continue;
                            var run = new StyleRun(row, edges[i + 1] - 1, fill, font, isBold, alignment);
                            if (runs.Count > 0 && runs[^1].To + 1 == run.From && runs[^1] with { From = run.From, To = run.To } == run) runs[^1] = runs[^1] with { To = run.To };
                            else runs.Add(run);
                        }
                        if (runs.Count > 0) style.Columns[c] = runs;
                    }
                }
                finally { Release(ws); }
            }
        }
        finally {
            if (active > 0) { dynamic ws = wb.Worksheets[active]; try { ws.Activate(); } catch (COMException) { } finally { Release(ws); } }
            LastErrors = budget.Errors; LastCalls = budget.Calls; LastSeconds = budget.Clock.Elapsed.TotalSeconds;
        }
        result.Partial = budget.Partial; return result;
    }
    // Excel can return either Boolean for mixed Hidden states. Mixed dimensions are the sentinel:
    // a hidden row/column has zero height/width, while visible rows/columns have positive dimensions.
    static object? Hidden(object target, bool rows, Budget budget)
    {
        dynamic range = target;
        object? hidden = range.Hidden;
        if (Mixed(hidden)) return null;
        if (!budget.Take()) return null;
        object? size = rows ? (object)range.RowHeight : (object)range.ColumnWidth;
        return Mixed(size) ? null : hidden;
    }

    sealed record PropRun(int From, int To, int Value);
    static bool Mixed(object? v) => v == null || v is DBNull;
    static List<PropRun> Probe(int from, int to, Func<int, int, object?> read, Budget budget)
    {
        var runs = new List<PropRun>();
        void Visit(int a, int b) {
            if (a > b || !budget.Take()) return;
            object? value;
            try { value = read(a, b); } catch (COMException) { budget.Errors++; budget.Partial = true; return; }
            if (!Mixed(value)) {
                int n = value is bool flag ? (flag ? 1 : 0) : Convert.ToInt32(value);
                if (runs.Count > 0 && runs[^1].To + 1 == a && runs[^1].Value == n) runs[^1] = runs[^1] with { To = b };
                else runs.Add(new PropRun(a, b, n));
            }
            else if (a < b) { int mid = a + (b - a) / 2; Visit(a, mid); Visit(mid + 1, b); }
            else budget.Partial = true;
        }
        Visit(from, to); return runs;
    }
    static void Release(object? obj) { if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
    public static StyleSidecar? Load(string source, CacheData cache)
    {
        string path = CachePolicy.SidecarPathFor(source); CachePolicy.Validate(path);
        try {
            if (!cache.IsFreshFor(source)) return null;
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var r = new BinaryReader(stream, Encoding.UTF8);
            if (Encoding.ASCII.GetString(r.ReadBytes(8)) != Magic || r.ReadInt64() != cache.SourceLength || r.ReadInt64() != cache.SourceTicks) return null;
            var result = new StyleSidecar { Partial = r.ReadBoolean() };
            int Count(int max) { int n = r.ReadInt32(); if (n < 0 || n > max || n > stream.Length - stream.Position) throw new InvalidDataException(); return n; }
            List<(int From, int To)> Ranges(int max) {
                int n = Count(max); var list = new List<(int, int)>(); int prev = -1;
                for (int i = 0; i < n; i++) { int a = r.ReadInt32(), b = r.ReadInt32(); if (a <= prev || b < a || b >= max) throw new InvalidDataException(); list.Add((a, b)); prev = b; } return list;
            }
            int sheets = Count(cache.Sheets.Count);
            for (int s = 0; s < sheets; s++) {
                int index = r.ReadInt32(); if (index < 0 || index >= cache.Sheets.Count) return null;
                var d = cache.Sheets[index]; var sheet = new SheetStyles { FrozenRows = r.ReadInt32(), FrozenCols = r.ReadInt32() };
                if (sheet.FrozenRows < 0 || sheet.FrozenRows >= 1048576 || sheet.FrozenCols < 0 || sheet.FrozenCols >= 16384) return null;
                sheet.HiddenRows = Ranges(d.Rows); sheet.HiddenCols = Ranges(d.Cols);
                int cols = Count(d.Cols);
                for (int c = 0; c < cols; c++) {
                    int col = r.ReadInt32(); if (col < 0 || col >= d.Cols) return null;
                    int n = Count(d.Rows), prev = -1; var runs = new List<StyleRun>();
                    for (int i = 0; i < n; i++) {
                        int a = r.ReadInt32(), b = r.ReadInt32(); if (a <= prev || b < a || b >= d.Rows) return null;
                        int? fill = ReadInt(r), font = ReadInt(r), bold = ReadInt(r), align = ReadInt(r);
                        if (fill is < 0 or > 0xffffff || font is < 0 or > 0xffffff || bold is < 0 or > 1 || align is < 0 or > 2) return null;
                        runs.Add(new StyleRun(a, b, fill, font, bold.HasValue ? bold == 1 : null, align)); prev = b;
                    }
                    sheet.Columns.Add(col, runs);
                }
                result.Sheets.Add(index, sheet);
            }
            return stream.Position == stream.Length ? result : null;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or OverflowException) { return null; }
    }
    public static StyleDelta? StyleAt(StyleSidecar sidecar, int sheet, int row, int col)
    {
        if (!sidecar.Sheets.TryGetValue(sheet, out var s) || !s.Columns.TryGetValue(col, out var runs)) return null;
        int lo = 0, hi = runs.Count - 1;
        while (lo <= hi) { int mid = lo + (hi - lo) / 2; var run = runs[mid]; if (row < run.From) hi = mid - 1; else if (row > run.To) lo = mid + 1; else return new StyleDelta(Fill: run.Fill, FontColor: run.FontColor, Bold: run.Bold, Align: run.Align); }
        return null;
    }
}
