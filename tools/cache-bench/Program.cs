// Tests the memory-first viewer design on made-up workbooks only.
// A separate reader process owns a hidden, read-only Excel and sends rows through a pipe.
// The viewer process keeps the rows in memory. Nothing is ever written to disk.
// Prints only timings and sizes, never cell contents.
//   dotnet cache-bench.dll viewer <syn*.xlsx> <columnar|boxed> [--abort-after-ms N]
// The approved disk cache is separate and OFF by default (see DiskCache.cs):
//   dotnet cache-bench.dll build-cache <file> <local folder> --disk-cache-approved
//   dotnet cache-bench.dll open-cache <cache file>
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

static class Sys
{
    [StructLayout(LayoutKind.Sequential)] struct MS { public uint len, load; public ulong tp, ap, tf, af, tv, av, ae; }
    [DllImport("kernel32.dll")] static extern bool GlobalMemoryStatusEx(ref MS m);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    public static double Commit() { var m = new MS { len = (uint)Marshal.SizeOf<MS>() }; GlobalMemoryStatusEx(ref m); return 100.0 * (m.tf - m.af) / m.tf; }
    public static long Priv(int id) { if (id == 0) return 0; try { using var p = Process.GetProcessById(id); return p.PrivateMemorySize64; } catch { return 0; } }
    public static long Self() { using var p = Process.GetCurrentProcess(); return p.PrivateMemorySize64; }
    public static string Col(int n) { var s = ""; while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; } return s; }
}

// ---------- The background reader: starts when needed, ends when done. ----------
static class Reader
{
    const int BlockRows = 1000;

    public static int Run(string path)
    {
        var sw = Stopwatch.StartNew();
        using var outS = new BufferedStream(Console.OpenStandardOutput(), 1 << 20);
        using var w = new BinaryWriter(outS, Encoding.UTF8);
        var jumps = new ConcurrentQueue<int>(); bool viewerGone = false;
        var input = new Thread(() =>
        {
            // When the viewer closes, this pipe ends, and the reader must stop too.
            string line; while ((line = Console.In.ReadLine()) != null) if (int.TryParse(line, out int b)) jumps.Enqueue(b);
            viewerGone = true;
        }) { IsBackground = true }; input.Start();

        dynamic xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        dynamic wb = null, ws = null, u = null; int stopCode = 0;
        try
        {
            xl.Visible = false; xl.DisplayAlerts = false; xl.AutomationSecurity = 3; xl.ScreenUpdating = false; xl.EnableEvents = false;
            Sys.GetWindowThreadProcessId(new IntPtr((long)xl.Hwnd), out uint pid);
            long startMs = sw.ElapsedMilliseconds;
            wb = xl.Workbooks.Open(path, 0, true); long openMs = sw.ElapsedMilliseconds - startMs;
            ws = wb.Worksheets[1]; u = ws.UsedRange;
            int rows = u.Rows.Count, cols = u.Columns.Count, blocks = (rows + BlockRows - 1) / BlockRows;
            w.Write(1); w.Write(rows); w.Write(cols); w.Write((int)pid); w.Write((int)startMs); w.Write((int)openMs); w.Flush();

            var done = new bool[blocks]; int next = 0, left = blocks; string last = Sys.Col(cols);
            while (left > 0 && !viewerGone)
            {
                double c = Sys.Commit();
                if (c >= 90) { stopCode = 90; break; }
                if (c >= 85) { w.Write(7); w.Flush(); Thread.Sleep(1000); continue; }
                int b = -1;
                while (jumps.TryDequeue(out int j)) if (j >= 0 && j < blocks && !done[j]) { b = j; break; }
                if (b < 0) { while (next < blocks && done[next]) next++; if (next >= blocks) break; b = next; }
                int r1 = b * BlockRows + 1, r2 = Math.Min(rows, r1 + BlockRows - 1), n = r2 - r1 + 1;
                var a = (object[,])ws.Range[$"A{r1}:{last}{r2}"].Value2;
                w.Write(2); w.Write(b); w.Write(n);
                for (int i = 1; i <= n; i++)
                    for (int k = 1; k <= cols; k++)
                        switch (a[i, k])
                        {
                            case null: w.Write((byte)0); break;
                            case double d: w.Write((byte)1); w.Write(d); break;
                            case string s: w.Write((byte)2); w.Write(s); break;
                            case bool t: w.Write((byte)3); w.Write(t); break;
                            default: w.Write((byte)4); w.Write(Convert.ToDouble(a[i, k])); break;
                        }
                w.Flush(); done[b] = true; left--;
            }
            w.Write(9); w.Write(stopCode); w.Flush();
        }
        catch (Exception e) { Console.Error.WriteLine("reader error: " + e.GetType().Name); stopCode = 1; }
        finally
        {
            // Close only the Excel this reader started. Never kill it.
            try { if (wb != null) wb.Close(false); } catch { }
            if (u != null) Marshal.FinalReleaseComObject(u); if (ws != null) Marshal.FinalReleaseComObject(ws); if (wb != null) Marshal.FinalReleaseComObject(wb);
            u = ws = wb = null;
            try { xl.Quit(); } catch { }
            Marshal.FinalReleaseComObject(xl); xl = null; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        }
        return stopCode;
    }
}

// ---------- The in-memory copies being compared. ----------
interface IStore
{
    void AddBlock(int idx, int n, int cols, BinaryReader br);
    string[] Screen(int row, int n);
    bool HasRow(int row);
    long SearchMiss(string needle);
    int[] SortNumber(int col);
    int[] SortText(int col);
}

// Baseline: keep what COM gives: one boxed object per cell.
sealed class BoxedStore : IStore
{
    readonly object[][,] blocks; readonly int cols, rows;
    public BoxedStore(int rows, int cols) { this.rows = rows; this.cols = cols; blocks = new object[(rows + 999) / 1000][,]; }
    public void AddBlock(int idx, int n, int cols, BinaryReader br)
    {
        var a = new object[n, cols];
        for (int i = 0; i < n; i++) for (int k = 0; k < cols; k++)
                switch (br.ReadByte()) { case 1: case 4: a[i, k] = br.ReadDouble(); break; case 2: a[i, k] = br.ReadString(); break; case 3: a[i, k] = br.ReadBoolean(); break; }
        Volatile.Write(ref blocks[idx], a);
    }
    public bool HasRow(int r) => Volatile.Read(ref blocks[r / 1000]) != null;
    object Get(int r, int c) => blocks[r / 1000][r % 1000, c];
    public string[] Screen(int row, int n)
    {
        var o = new string[n * cols]; int p = 0;
        for (int r = row; r < row + n; r++) for (int c = 0; c < cols; c++) o[p++] = Get(r, c)?.ToString() ?? "";
        return o;
    }
    public long SearchMiss(string needle)
    {
        long hits = 0;
        foreach (var b in blocks) { if (b == null) continue; foreach (var v in b) if (v is string s && s.Contains(needle, StringComparison.Ordinal)) hits++; }
        return hits;
    }
    public int[] SortNumber(int col)
    {
        var k = new double[rows]; var ix = new int[rows];
        for (int r = 0; r < rows; r++) { ix[r] = r; k[r] = Get(r, col) is double d ? d : double.NaN; }
        Array.Sort(k, ix); return ix;
    }
    public int[] SortText(int col)
    {
        var ix = new int[rows]; for (int r = 0; r < rows; r++) ix[r] = r;
        Array.Sort(ix, (a, b) => string.CompareOrdinal(Get(a, col) as string, Get(b, col) as string)); return ix;
    }
}

// Our copy: per 1,000-row chunk, per column: plain numbers, or text bytes packed together.
// Repeated text is stored once. A kinds array exists only when a column mixes kinds in that chunk.
sealed class ColumnChunk
{
    public byte[] Kinds;      // null when every cell has the same kind (SolidKind)
    public byte SolidKind;
    public double[] Nums;
    public byte[] Blob;       // text bytes
    public int[] Offs;        // start of each text; one extra at the end
    public ushort[] Index;    // when repeated text is shared: row -> entry in Blob/Offs
    public byte KindAt(int r) => Kinds == null ? SolidKind : Kinds[r];
    public ReadOnlySpan<byte> TextAt(int r) { int e = Index != null ? Index[r] : r; return Blob.AsSpan(Offs[e], Offs[e + 1] - Offs[e]); }
    public string Display(int r) => KindAt(r) switch
    {
        0 => "",
        2 => Encoding.UTF8.GetString(TextAt(r)),
        3 => Nums[r] != 0 ? "TRUE" : "FALSE",
        4 => "#ERR",
        _ => Nums[r].ToString("R"),
    };
}

sealed class ColumnarStore : IStore
{
    readonly ColumnChunk[][] chunks; readonly int cols, rows;
    public ColumnarStore(int rows, int cols) { this.rows = rows; this.cols = cols; chunks = new ColumnChunk[(rows + 999) / 1000][]; }

    public void AddBlock(int idx, int n, int cols, BinaryReader br)
    {
        var kinds = new byte[cols][]; var nums = new double[cols][]; var text = new string[cols][];
        for (int i = 0; i < n; i++) for (int k = 0; k < cols; k++)
            {
                kinds[k] ??= new byte[n];
                byte t = br.ReadByte(); kinds[k][i] = t;
                if (t == 2) (text[k] ??= new string[n])[i] = br.ReadString();
                else if (t == 1 || t == 4) (nums[k] ??= new double[n])[i] = br.ReadDouble();
                else if (t == 3) (nums[k] ??= new double[n])[i] = br.ReadBoolean() ? 1 : 0;
            }
        var cc = new ColumnChunk[cols];
        for (int k = 0; k < cols; k++) cc[k] = Build(n, kinds[k] ?? new byte[n], nums[k], text[k]);
        Volatile.Write(ref chunks[idx], cc); // ready and never changed again, so the screen can read it with no lock
    }

    internal static ColumnChunk Build(int n, byte[] kinds, double[] nums, string[] text)
    {
        var c = new ColumnChunk { Nums = nums };
        byte first = kinds[0]; bool same = true; for (int i = 1; i < n; i++) if (kinds[i] != first) { same = false; break; }
        if (same) c.SolidKind = first; else c.Kinds = kinds;
        if (text == null) return c;
        var map = new Dictionary<string, ushort>(); var list = new List<string>(); var index = new ushort[n]; bool share = true;
        for (int i = 0; i < n; i++)
        {
            string s = text[i] ?? "";
            if (!map.TryGetValue(s, out ushort e)) { if (list.Count >= n / 2) { share = false; break; } e = (ushort)list.Count; map[s] = e; list.Add(s); }
            index[i] = e;
        }
        IList<string> entries = share ? list : Array.ConvertAll(text, s => s ?? "");
        var offs = new int[entries.Count + 1]; int size = 0;
        for (int i = 0; i < entries.Count; i++) { offs[i] = size; size += Encoding.UTF8.GetByteCount(entries[i]); }
        offs[entries.Count] = size;
        var blob = new byte[size];
        for (int i = 0; i < entries.Count; i++) Encoding.UTF8.GetBytes(entries[i], 0, entries[i].Length, blob, offs[i]);
        c.Blob = blob; c.Offs = offs; if (share) c.Index = index;
        return c;
    }

    public bool HasRow(int r) => Volatile.Read(ref chunks[r / 1000]) != null;
    public string[] Screen(int row, int n)
    {
        var o = new string[n * cols]; int p = 0;
        for (int r = row; r < row + n; r++) { var cc = chunks[r / 1000]; int i = r % 1000; for (int c = 0; c < cols; c++) o[p++] = cc[c].Display(i); }
        return o;
    }
    public long SearchMiss(string needle)
    {
        var nb = Encoding.UTF8.GetBytes(needle); long hits = 0;
        foreach (var cc in chunks) { if (cc == null) continue; foreach (var c in cc) if (c.Blob != null && c.Blob.AsSpan().IndexOf(nb) >= 0) hits++; }
        return hits;
    }
    public int[] SortNumber(int col)
    {
        var k = new double[rows]; var ix = new int[rows];
        for (int r = 0; r < rows; r++) { ix[r] = r; var c = chunks[r / 1000][col]; int i = r % 1000; k[r] = c.KindAt(i) == 1 ? c.Nums[i] : double.NaN; }
        Array.Sort(k, ix); return ix;
    }
    public int[] SortText(int col)
    {
        var ix = new int[rows]; for (int r = 0; r < rows; r++) ix[r] = r;
        Array.Sort(ix, (a, b) => chunks[a / 1000][col].TextAt(a % 1000).SequenceCompareTo(chunks[b / 1000][col].TextAt(b % 1000)));
        return ix;
    }
}

// ---------- The viewer side of the test. ----------
static class Viewer
{
    static string Med(List<double> t) { if (t.Count == 0) return "n=0"; t.Sort(); return $"n={t.Count} median={t[t.Count / 2]:F2}ms p95={t[(int)(t.Count * 0.95)]:F2}ms max={t[^1]:F2}ms"; }

    public static int Run(string path, string storeName, int abortMs)
    {
        if (!Path.GetFileName(path).StartsWith("syn", StringComparison.OrdinalIgnoreCase)) { Console.WriteLine("REFUSED: this benchmark is for made-up files only"); return 2; }
        for (int i = 0; i < 300 && Sys.Commit() >= 85; i++) { if (i == 0) Console.WriteLine("waiting: commit >= 85%"); Thread.Sleep(1000); }
        if (Sys.Commit() >= 85) { Console.WriteLine("SKIPPED: commit stayed >= 85%"); return 3; }
        Console.WriteLine($"store={storeName} commit at start {Sys.Commit():F1}% viewer={Sys.Self() / 1048576}MB");
        var all = Stopwatch.StartNew();
        var psi = new ProcessStartInfo("dotnet") { RedirectStandardOutput = true, RedirectStandardInput = true, UseShellExecute = false, CreateNoWindow = true };
        psi.ArgumentList.Add(Environment.GetCommandLineArgs()[0]); psi.ArgumentList.Add("reader"); psi.ArgumentList.Add(path);
        using var child = Process.Start(psi);
        var br = new BinaryReader(child.StandardOutput.BaseStream, Encoding.UTF8);

        if (br.ReadInt32() != 1) { Console.WriteLine("reader failed before opening"); return 4; }
        int rows = br.ReadInt32(), cols = br.ReadInt32(), xpid = br.ReadInt32(), startMs = br.ReadInt32(), openMs = br.ReadInt32();
        double openedAt = all.Elapsed.TotalMilliseconds;
        Console.WriteLine($"{rows}x{cols}; excel start {startMs}ms, open {openMs}ms; row count known at {openedAt:F0}ms");
        IStore store = storeName == "boxed" ? new BoxedStore(rows, cols) : new ColumnarStore(rows, cols);
        int blocks = (rows + 999) / 1000;

        long peakExcel = 0, peakChild = 0, peakSelf = 0; double peakCommit = 0; bool sampling = true; int cpid = child.Id;
        var sampler = new Thread(() => { while (Volatile.Read(ref sampling)) { peakExcel = Math.Max(peakExcel, Sys.Priv(xpid)); peakChild = Math.Max(peakChild, Sys.Priv(cpid)); peakSelf = Math.Max(peakSelf, Sys.Self()); peakCommit = Math.Max(peakCommit, Sys.Commit()); Thread.Sleep(250); } }) { IsBackground = true };
        sampler.Start();

        // A person jumping around while the rest still loads.
        var pending = new ConcurrentDictionary<int, long>(); var jumpT = new List<double>(); bool loading = true; var rng = new Random(5);
        var jumper = new Thread(() =>
        {
            Thread.Sleep(300);
            for (int i = 0; i < 40 && Volatile.Read(ref loading); i++)
            {
                int b = rng.Next(blocks);
                if (!store.HasRow(b * 1000) && pending.TryAdd(b, Stopwatch.GetTimestamp())) { lock (child) { child.StandardInput.WriteLine(b); child.StandardInput.Flush(); } }
                Thread.Sleep(300);
            }
        }) { IsBackground = true };

        if (abortMs > 0)
        {
            new Thread(() =>
            {
                Thread.Sleep(abortMs); lock (child) { child.StandardInput.Close(); }
                Console.WriteLine($"viewer closed the pipe at {all.ElapsedMilliseconds}ms");
            }) { IsBackground = true }.Start();
        }
        else jumper.Start();

        double firstScreen = -1, screenOnly = -1; int got = 0, stop = 0; double waitFrames = 0; var decode = new Stopwatch();
        while (true)
        {
            int t;
            try { t = br.ReadInt32(); } catch (EndOfStreamException) { break; }
            if (t == 7) { waitFrames++; continue; }
            if (t == 9) { stop = br.ReadInt32(); break; }
            int b = br.ReadInt32(), n = br.ReadInt32();
            decode.Start(); store.AddBlock(b, n, cols, br); decode.Stop(); got++;
            if (firstScreen < 0)
            {
                var s = Stopwatch.StartNew(); store.Screen(0, Math.Min(40, n)); screenOnly = s.Elapsed.TotalMilliseconds;
                firstScreen = all.Elapsed.TotalMilliseconds;
                Console.WriteLine($"FIRST SCREEN {firstScreen:F0}ms from start ({firstScreen - openedAt:F1}ms after Excel finished opening; drawing data from store {screenOnly:F3}ms)");
            }
            if (pending.TryRemove(b, out long st))
            {
                store.Screen(b * 1000, Math.Min(40, n));
                jumpT.Add((Stopwatch.GetTimestamp() - st) * 1000.0 / Stopwatch.Frequency);
            }
        }
        double fullRead = all.Elapsed.TotalMilliseconds; Volatile.Write(ref loading, false);
        Console.WriteLine($"FULL READ {fullRead / 1000:F2}s ({got}/{blocks} blocks, {got * 1000.0 * cols / ((fullRead - openedAt) / 1000) / 1e6:F2}M cells/s after open); decode+store {decode.Elapsed.TotalSeconds:F2}s; stop code {stop}; paused {waitFrames}s for commit");
        Console.WriteLine($"jump to an unread place while loading: {Med(jumpT)}");
        double beforeClose = Sys.Commit(); long excelAtEnd = Sys.Priv(xpid);

        child.StandardInput.Close();
        var cl = Stopwatch.StartNew();
        while (Sys.Priv(xpid) != 0 && cl.Elapsed.TotalSeconds < 60) Thread.Sleep(100);
        Console.WriteLine($"reader stopped; own Excel {(Sys.Priv(xpid) == 0 ? "left" : "STILL THERE, not killed")} after {cl.Elapsed.TotalSeconds:F1}s");
        Volatile.Write(ref sampling, false); sampler.Join();
        Console.WriteLine($"peaks during load: excel={peakExcel / 1048576}MB reader={peakChild / 1048576}MB viewer={peakSelf / 1048576}MB commit={peakCommit:F1}%");
        if (abortMs > 0) { Console.WriteLine($"commit after close {Sys.Commit():F1}%"); return 0; }

        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        Console.WriteLine($"MEMORY after load, Excel gone: viewer={Sys.Self() / 1048576}MB (live data {GC.GetTotalMemory(true) / 1048576}MB; ~{(double)GC.GetTotalMemory(false) / ((long)rows * cols):F1} bytes/cell) vs Excel held open {excelAtEnd / 1048576}MB; commit {beforeClose:F1}% -> {Sys.Commit():F1}%");

        var r2 = new Random(9); var sc = new List<double>();
        for (int i = 0; i < 2000; i++) { int r = r2.Next(0, Math.Max(1, rows - 40)); var s = Stopwatch.StartNew(); store.Screen(r, 40); sc.Add(s.Elapsed.TotalMilliseconds); }
        Console.WriteLine($"screen of 40 rows from memory: {Med(sc)}");
        var sw = Stopwatch.StartNew(); long hits = store.SearchMiss("NO-SUCH-VALUE-ZZZ"); Console.WriteLine($"search all text, nothing found: {sw.ElapsedMilliseconds}ms hits={hits}");
        sw.Restart(); store.SortNumber(0); Console.WriteLine($"sort by a number column: {sw.ElapsedMilliseconds}ms");
        sw.Restart(); store.SortText(5); Console.WriteLine($"sort by a repeated-text column: {sw.ElapsedMilliseconds}ms");
        sw.Restart(); store.SortText(8); Console.WriteLine($"sort by a unique-text column: {sw.ElapsedMilliseconds}ms");
        Console.WriteLine($"viewer after sorts: {Sys.Self() / 1048576}MB; commit {Sys.Commit():F1}%");
        GC.KeepAlive(store);
        return 0;
    }
}

static class P
{
    [STAThread]
    static int Main(string[] a)
    {
        if (a.Length < 2) { Console.WriteLine("usage: viewer <syn*.xlsx> <columnar|boxed> [--abort-after-ms N]"); return 2; }
        if (a[0] == "reader") return Reader.Run(a[1]);
        if (a[0] == "build-cache") return Build.Run(a[1], a[2], a.Contains("--disk-cache-approved"), a.Contains("--stop-only-at-90"));
        if (a[0] == "open-cache") return Reopen.Run(a[1]);
        int abort = 0; int ix = Array.IndexOf(a, "--abort-after-ms"); if (ix >= 0) abort = int.Parse(a[ix + 1]);
        return Viewer.Run(a[1], a.Length > 2 ? a[2] : "columnar", abort);
    }
}
