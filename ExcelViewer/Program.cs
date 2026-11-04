using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Text;
using ExcelViewer.Core;

namespace ExcelViewer;

static class Program
{
    [STAThread]
    static int Main(string[] a)
    {
        if (a.Length >= 2 && a[0] == "--build") return BuildMode(a[1]);
        ApplicationConfiguration.Initialize();
        bool unattended = a.Length >= 1 && a[0] is "--timing" or "--shot";
        // Never show an error message with file contents in it, and never leave a test run waiting on a box.
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (s, e) =>
        {
            if (unattended) { Console.WriteLine("ERROR " + e.Exception.GetType().Name); Environment.Exit(1); }
            MessageBox.Show("Something went wrong (" + e.Exception.GetType().Name + "). The source file is not changed.", "Excel Viewer");
        };
        if (a.Length >= 2 && a[0] == "--timing") return TimingMode(a[1]);
        if (a.Length >= 3 && a[0] == "--shot") return ShotMode(a[1], a[2], a.Length > 3 ? a[3] : null, a.Length > 4 ? int.Parse(a[4]) : -1);
        var f = new MainForm();
        if (a.Length >= 1 && File.Exists(a[0])) f.Shown += (s, e) => f.OpenFile(a[0]);
        Application.Run(f);
        return 0;
    }

    // The reader. It owns the hidden Excel. If the viewer closes, its pipe ends and the reader stops too.
    static int BuildMode(string source)
    {
        using var o = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
        bool gone = false;
        new Thread(() => { try { while (Console.In.ReadLine() != null) { } } catch (IOException) { } gone = true; }) { IsBackground = true }.Start();
        var r = CacheBuilder.Build(source, CachePolicy.PathFor(source), o.WriteLine, () => gone);
        o.WriteLine(string.Create(CultureInfo.InvariantCulture, $"RESULT code={r.Code} stopped={r.Stopped} excelStart={r.ExcelStartSec:F1}s open={r.OpenSec:F1}s formats={r.FormatSec:F1}s read={r.ReadSec:F1}s check={r.VerifySec:F1}s sheets={r.Sheets} cells={r.Cells} cacheMB={r.CacheBytes / 1048576.0:F1} formatsTotal={r.Formats} formatsMatched={r.FormatsOk} formatsNotCarried={r.FormatsFailed} formatsNoCells={r.FormatsUnused} unreadFormatRuns={r.UnreadRuns} peakExcelMB={r.PeakExcelMB} peakCommit={r.PeakCommit:F1}%"));
        return r.Code;
    }

    // Times the real window, off screen, on an existing copy. Prints only times and counts.
    static int TimingMode(string source)
    {
        string cp = CachePolicy.PathFor(source);
        if (!CacheData.HeaderMatches(cp, source)) { Console.WriteLine("NO FRESH COPY"); return 3; }
        double fromStart = (DateTime.Now - Process.GetCurrentProcess().StartTime).TotalMilliseconds;
        var form = new MainForm { Unattended = true }; form.PlaceOffScreen(); int code = 0;
        form.Shown += (s, e) =>
        {
            try
            {
                var sw = Stopwatch.StartNew();
                form.Source(source); form.LoadForTest(cp);
                var bmp = new Bitmap(form.Grid.Width, form.Grid.Height);
                form.Grid.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                double total = (DateTime.Now - Process.GetCurrentProcess().StartTime).TotalMilliseconds;
                Console.WriteLine($"FIRST SCREEN drawn: {total:F0} ms from program start (program ready at {fromStart:F0} ms; copy to screen {sw.Elapsed.TotalMilliseconds:F0} ms)");
                for (int i = 0; i < 3; i++) { sw.Restart(); form.OpenFile(source); form.Grid.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height)); Console.WriteLine($"OPEN AGAIN in the running window: {sw.Elapsed.TotalMilliseconds:F0} ms"); }
                var c = form.Cache; int largest = 0; for (int i = 0; i < c.Sheets.Count; i++) if (c.Sheets[i].Cells > c.Sheets[largest].Cells) largest = i;
                sw.Restart(); form.ShowSheet(largest); form.Grid.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                Console.WriteLine($"SWITCH to the largest sheet: {sw.Elapsed.TotalMilliseconds:F0} ms");
                var d = form.Reader.Sheet; var g = form.Grid;
                Console.WriteLine($"largest sheet: {d.Rows} rows x {d.Cols} columns; memory {Mem()} MB");

                var rng = new Random(3); var t = new List<double>();
                for (int i = 0; i < 200; i++) { sw.Restart(); g.FirstDisplayedScrollingRowIndex = rng.Next(0, Math.Max(1, g.RowCount - 40)); g.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height)); t.Add(sw.Elapsed.TotalMilliseconds); }
                t.Sort(); Console.WriteLine($"SCROLL long jump + full redraw: median {t[100]:F1} ms, p95 {t[190]:F1} ms, worst {t[199]:F1} ms; memory {Mem()} MB");
                t.Clear(); int row = 0;
                for (int i = 0; i < 200; i++) { sw.Restart(); row = Math.Min(g.RowCount - 1, row + 3); g.FirstDisplayedScrollingRowIndex = row; g.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height)); t.Add(sw.Elapsed.TotalMilliseconds); }
                t.Sort(); Console.WriteLine($"SCROLL mouse-wheel step + full redraw: median {t[100]:F1} ms, p95 {t[190]:F1} ms, worst {t[199]:F1} ms");

                var all = Enumerable.Range(0, c.Sheets.Count).ToList(); long n = 0;
                sw.Restart(); SearchEngine.Run(c, all, "NO-SUCH-VALUE-ZZZ", _ => n++, _ => { }, CancellationToken.None);
                Console.WriteLine($"SEARCH words, every sheet, nothing found: {sw.Elapsed.TotalSeconds:F2} s; hits {n}; memory {Mem()} MB");
                n = 0; sw.Restart(); SearchEngine.Run(c, all, "98765432100123", _ => n++, _ => { }, CancellationToken.None);
                Console.WriteLine($"SEARCH digits (checks every shown number), every sheet, nothing found: {sw.Elapsed.TotalSeconds:F2} s; hits {n}; memory {Mem()} MB");

                int nc = -1, tc = -1; var rd = form.Reader;
                for (int k = 0; k < d.Cols && (nc < 0 || tc < 0); k++) { byte kd = rd.Column(Math.Min(d.Rows - 1, 1), k).KindAt(Math.Min(d.Rows - 1, 1) % d.BlockRows); if (kd == CellKind.Number && nc < 0) nc = k; if (kd == CellKind.Text && tc < 0) tc = k; }
                if (nc >= 0) { sw.Restart(); var m = SortEngine.Sort(c, d, nc, false, form.FirstDataRow, CancellationToken.None); double st = sw.Elapsed.TotalMilliseconds; form.RowMap = m; form.Redraw(); sw.Restart(); g.FirstDisplayedScrollingRowIndex = 0; g.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height)); Console.WriteLine($"SORT a number column: {st:F0} ms; redraw {sw.Elapsed.TotalMilliseconds:F0} ms"); }
                if (tc >= 0) { sw.Restart(); SortEngine.Sort(c, d, tc, false, form.FirstDataRow, CancellationToken.None); Console.WriteLine($"SORT a text column: {sw.Elapsed.TotalMilliseconds:F0} ms; memory {Mem()} MB"); }

                sw.Restart(); bool same = c.Sheets.All(x => { using var r = new SheetReader(c, x); return r.RecomputeHash() == x.Hash; });
                Console.WriteLine($"copy reads back identical: {(same ? "YES" : "NO")} ({sw.Elapsed.TotalSeconds:F1} s)");
                foreach (var (x, i) in c.Sheets.Select((x, i) => (x, i))) { using var r = new SheetReader(c, x); int w = r.ColumnsWithFormatWarnings(); if (w > 0) Console.WriteLine($"sheet#{i + 1}: {w} of {x.Cols} columns have a format not carried"); }
                Console.WriteLine($"formats: {c.Formats.Count} total, {c.Formats.Count(f => f.Status == FormatStatus.MatchesExcel)} match Excel, {c.Formats.Count(f => f.Status == FormatStatus.CannotCarry)} not carried, {c.Formats.Count(f => f.Status == FormatStatus.NotChecked)} on no cell");
            }
            catch (Exception ex) { Console.WriteLine("ERROR " + ex.GetType().Name); code = 1; }
            form.Close();
        };
        Application.Run(form);
        return code;
    }

    static long Mem() { using var p = Process.GetCurrentProcess(); return p.PrivateMemorySize64 / 1048576; }

    // Drives the real window through its real steps, off screen, and takes pictures. Made-up files only.
    static int ShotMode(string source, string png, string search, int sortCol)
    {
        if (!Path.GetFileName(source).StartsWith("syn", StringComparison.OrdinalIgnoreCase)) { Console.WriteLine("REFUSED: pictures are for made-up files only"); return 2; }
        var form = new MainForm { Unattended = true }; form.PlaceOffScreen(); form.ClientSize = new Size(1180, 640); int code = 0, state = 0;
        var clock = Stopwatch.StartNew(); var timer = new System.Windows.Forms.Timer { Interval = 100 };
        void Save(string file) { using var b = new Bitmap(form.Width, form.Height); form.DrawToBitmap(b, new Rectangle(0, 0, b.Width, b.Height)); b.Save(file); }
        timer.Tick += (s, e) =>
        {
            if (clock.Elapsed.TotalSeconds > 180) { Console.WriteLine("TIMED OUT in step " + state); code = 1; timer.Stop(); form.Close(); return; }
            if (state == 0 && form.Cache != null)
            {
                Console.WriteLine($"opened in {clock.Elapsed.TotalSeconds:F2} s; rows on screen list {form.Grid.RowCount}");
                Save(Path.ChangeExtension(png, ".open.png"));
                if (sortCol >= 0) form.ClickSort(sortCol);
                state = 1;
            }
            else if (state == 1 && (sortCol < 0 || form.SortedColumn == sortCol))
            {
                Console.WriteLine("sort: " + form.InfoText);
                if (!string.IsNullOrEmpty(search)) form.SearchFor(search);
                state = 2;
            }
            else if (state == 2 && (string.IsNullOrEmpty(search) || form.SearchFinished))
            {
                Console.WriteLine($"search: {form.InfoText}; hits {form.HitCount}");
                form.JumpToFirstHit(); Save(png);
                timer.Stop(); form.Close();
            }
        };
        form.Shown += (s, e) => { form.OpenFile(source); timer.Start(); };
        Application.Run(form);
        return code;
    }
}
