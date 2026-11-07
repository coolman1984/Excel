// Excel owns workbook writes. No workbook parser or writer is used here.
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelViewer.Core;

namespace ExcelApp;

public static class SaveBack
{
    public sealed class Job
    {
        public string Source { get; set; } = "";
        public string Target { get; set; } = "";
        public bool SameFile { get; set; }
        public long? ExpectedLength { get; set; }
        public long? ExpectedTicks { get; set; }
        public List<Edit> Edits { get; set; } = new();
        public List<Width> ColWidths { get; set; } = new();
        public List<Freeze> Freeze { get; set; } = new();
    }
    public sealed class Edit
    {
        public string Sheet { get; set; } = "";
        public int Row { get; set; }
        public int Col { get; set; }
        public EditKind Kind { get; set; }
        public double Num { get; set; }
        public string? Text { get; set; }
        public string? Formula { get; set; }
        public bool? Bool { get; set; }
        public StyleDelta? Style { get; set; }
        // A style-only edit must not clear an existing value.
        public bool StyleOnly { get; set; }
    }
    public sealed class Width { public string Sheet { get; set; } = ""; public int Col { get; set; } public double Px { get; set; } }
    public sealed class Freeze { public string Sheet { get; set; } = ""; public int Rows { get; set; } public int Cols { get; set; } }

    public static int Run(TextReader stdin, TextWriter stdout, double? stopAt = null)
    {
        var timer = Stopwatch.StartNew(); int code = 1, count = 0;
        dynamic? xl = null, wb = null;
        var limit = Math.Min(stopAt ?? Machine.StopAt, Machine.StopAt);
        bool cancelled = false;
        void Line(string s) { lock (stdout) { stdout.WriteLine(s); stdout.Flush(); } }
        void Check() {
            if (Machine.Commit() >= limit) throw new SaveStop(90, "memory");
            if (Volatile.Read(ref cancelled)) throw new SaveStop(2, "cancelled");
        }
        try
        {
            Check();
            // One JSON line followed by an open stdin pipe; EOF afterwards requests cancellation.
            var line = stdin.ReadLine() ?? throw new InvalidDataException();
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            options.Converters.Add(new JsonStringEnumConverter());
            var job = JsonSerializer.Deserialize<Job>(line, options) ?? throw new InvalidDataException();
            Validate(job);
            void Fresh() {
                if (!job.SameFile || job.ExpectedLength == null || job.ExpectedTicks == null) return;
                var source = new FileInfo(job.Source);
                if (!source.Exists || source.Length != job.ExpectedLength || source.LastWriteTimeUtc.Ticks != job.ExpectedTicks) throw new SaveStop(7, "changed");
            }
            Fresh();
            if (!ReferenceEquals(stdin, Console.In)) { /* In-memory callers need no pipe watcher. */ }
            else _ = Task.Run(() => { try { while (stdin.Read() >= 0) { } } catch (IOException) { } Volatile.Write(ref cancelled, true); });
            Check(); Line("OPENING");
            xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!) ?? throw new InvalidOperationException();
            xl.Visible = false; xl.DisplayAlerts = false; xl.AutomationSecurity = 3;
            xl.ScreenUpdating = false; xl.EnableEvents = false; xl.Interactive = false;
            wb = xl.Workbooks.Open(job.Source, UpdateLinks: 0, ReadOnly: false, Notify: false, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true);
            Line("OPENED"); Check();
            if ((bool)wb.ReadOnly && job.SameFile) throw new SaveStop(4, "readonly");
            // Check every affected sheet before applying any changes.
            var names = job.Edits.Select(e => e.Sheet).Concat(job.ColWidths.Select(e => e.Sheet)).Concat(job.Freeze.Select(e => e.Sheet)).Distinct().ToArray();
            foreach (string name in names) { Check(); dynamic ws = wb.Worksheets[name]; try { if ((bool)ws.ProtectContents) throw new SaveStop(6, "protected"); } finally { Release(ws); } }
            int total = job.Edits.Count + job.ColWidths.Count + job.Freeze.Count, done = 0, last = -2;
            void Progress(int n) { done += n; int pct = total == 0 ? 100 : done * 100 / total; if (pct >= last + 2 || pct == 100) { Line("WRITING " + pct); last = pct; } }
            foreach (var group in job.Edits.GroupBy(e => e.Sheet))
            {
                dynamic ws = wb.Worksheets[group.Key];
                try
                {
                    // Same-kind adjacent cells on one row share one COM value write.
                    var edits = group.OrderBy(e => e.Row).ThenBy(e => e.Col).ToArray();
                    for (int i = 0; i < edits.Length;)
                    {
                        Check(); var first = edits[i]; int end = i + 1;
                        while (end < edits.Length && end - i < 1000 && edits[end].Row == first.Row && edits[end].Col == edits[end - 1].Col + 1 && edits[end].Kind == first.Kind && edits[end].StyleOnly == first.StyleOnly) end++;
                        dynamic range = ws.Range[$"{CacheBuilder.Letters(first.Col)}{first.Row}:{CacheBuilder.Letters(edits[end - 1].Col)}{first.Row}"];
                        try
                        {
                            if (!first.StyleOnly)
                            {
                                if (first.Kind == EditKind.Clear) range.ClearContents();
                                else {
                                    var values = new object[1, end - i];
                                    for (int j = i; j < end; j++) values[0, j - i] = Value(edits[j]);
                                    if (first.Kind == EditKind.Formula) range.Formula = values; else range.Value2 = values;
                                }
                            }
                            for (int j = i; j < end; j++) if (edits[j].Style != null) {
                                dynamic cell = ws.Cells[edits[j].Row, edits[j].Col];
                                try { WriteStyle((object)cell, edits[j].Style!); } finally { Release(cell); }
                            }
                        }
                        finally { Release(range); }
                        Progress(end - i); i = end;
                    }
                }
                finally { Release(ws); }
            }
            foreach (var width in job.ColWidths) { Check(); dynamic ws = wb.Worksheets[width.Sheet]; dynamic column = ws.Columns[width.Col]; try { column.ColumnWidth = Math.Max(0, (width.Px - 5) / 7); } finally { Release(column); Release(ws); } Progress(1); }
            foreach (var freeze in job.Freeze) {
                Check(); dynamic ws = wb.Worksheets[freeze.Sheet]; dynamic window = wb.Windows[1];
                try { ws.Activate(); window.FreezePanes = false; window.SplitRow = freeze.Rows; window.SplitColumn = freeze.Cols; window.FreezePanes = freeze.Rows > 0 || freeze.Cols > 0; }
                finally { Release(window); Release(ws); } Progress(1);
            }
            Check(); if ((int)xl.Calculation == -4135) xl.CalculateFull();
            Line("SAVING");
            if (job.SameFile) { Fresh(); wb.Save(); }
            else wb.SaveAs(job.Target, FileFormat: Format(job.Target), ConflictResolution: 2);
            Line("SAVED"); Check();
            var result = CacheBuilder.BuildFromOpenWorkbook((object)xl, (object)wb, job.Target, CachePolicy.PathFor(job.Target), Line,
                () => Volatile.Read(ref cancelled) || Machine.Commit() >= limit, StyleStore.AfterRead);
            code = result.Code; count = job.Edits.Count;
            if (result.Code != 0) Line("STOPPED " + (result.Stopped ?? "readback"));
        }
        catch (SaveStop e) { code = e.Code; Line("STOPPED " + e.Reason); }
        catch (Exception e) { code = 1; Line("STOPPED " + e.GetType().Name); }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch (Exception e) { Line("CLEANUP " + e.GetType().Name); if (code == 0) code = 1; }
            try { if (xl != null) xl.Quit(); } catch (Exception e) { Line("CLEANUP " + e.GetType().Name); if (code == 0) code = 1; }
            Release(wb); Release(xl);
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            Line(FormattableString.Invariant($"RESULT code={code} seconds={timer.Elapsed.TotalSeconds:F3} edits={count}"));
        }
        return code;
    }
    sealed class SaveStop(int code, string reason) : Exception { public int Code = code; public string Reason = reason; }
    static void Release(object? value) { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
    static int Format(string path) => Path.GetExtension(path).ToLowerInvariant() switch { ".xlsx" => 51, ".xlsm" => 52, ".xlsb" => 50, ".xls" => 56, _ => throw new InvalidDataException() };
    static void Validate(Job j)
    {
        if (string.IsNullOrWhiteSpace(j.Source) || string.IsNullOrWhiteSpace(j.Target)) throw new InvalidDataException();
        bool same = string.Equals(Path.GetFullPath(j.Source), Path.GetFullPath(j.Target), StringComparison.OrdinalIgnoreCase);
        if (same != j.SameFile) throw new InvalidDataException();
        _ = Format(j.Target);
        if (j.Edits.Any(e => e.Row < 1 || e.Row > 1048576 || e.Col < 1 || e.Col > 16384 || !Enum.IsDefined(e.Kind) || e.Kind == EditKind.Number && !double.IsFinite(e.Num))) throw new InvalidDataException();
        if (j.ColWidths.Any(e => e.Col < 1 || e.Col > 16384 || !double.IsFinite(e.Px) || e.Px < 5 || e.Px > 1790)) throw new InvalidDataException();
        if (j.Freeze.Any(e => e.Rows < 0 || e.Rows >= 1048576 || e.Cols < 0 || e.Cols >= 16384)) throw new InvalidDataException();
    }
    static object Value(Edit e) => e.Kind switch {
        EditKind.Number => e.Num, EditKind.Bool => e.Bool ?? e.Num != 0,
        EditKind.Formula => e.Formula ?? throw new InvalidDataException(),
        EditKind.Text => ProtectText(e.Text ?? ""), _ => throw new InvalidDataException()
    };
    static string ProtectText(string s) => s.Length > 0 && "=+-@'".Contains(s[0]) ? "'" + s : s;
    internal static int Bgr(int rgb) => ((rgb & 255) << 16) | (rgb & 0xff00) | ((rgb >> 16) & 255);
    static void WriteStyle(object cell, StyleDelta s)
    {
        dynamic c = cell;
        if (s.Bold.HasValue) c.Font.Bold = s.Bold.Value;
        if (s.Italic.HasValue) c.Font.Italic = s.Italic.Value;
        if (s.Underline.HasValue) c.Font.Underline = s.Underline.Value ? 2 : -4142;
        if (s.FontColor.HasValue) c.Font.Color = Bgr(s.FontColor.Value);
        if (s.Fill.HasValue) c.Interior.Color = Bgr(s.Fill.Value);
        if (s.Align.HasValue) c.HorizontalAlignment = s.Align.Value switch { 0 => -4131, 1 => -4152, 2 => -4108, _ => throw new InvalidDataException() };
        if (s.NumberFormat != null) c.NumberFormat = s.NumberFormat;
        if (s.Border != null) {
            if (s.Border is not ("none" or "all" or "outside" or "bottom")) throw new InvalidDataException();
            foreach (int edge in new[] { 7, 8, 9, 10, 11, 12 }) c.Borders[edge].LineStyle = s.Border == "all" || s.Border == "outside" && edge <= 10 || s.Border == "bottom" && edge == 9 ? 1 : -4142;
        }
    }
}
