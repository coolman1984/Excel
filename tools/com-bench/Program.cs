// Speed test for Excel COM. Prints only timings and sizes, never cell contents.
// "make" builds a made-up workbook. "bench" measures reading it.
// The fast-copy step runs only with --cache and only on files whose name starts with "syn" (made-up data).
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

static class P
{
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    [StructLayout(LayoutKind.Sequential)] struct MS { public uint len, load; public ulong tp, ap, tf, af, tv, av, ae; }
    [DllImport("kernel32.dll")] static extern bool GlobalMemoryStatusEx(ref MS m);
    static volatile bool stop;
    static long peakExcel, peakSelf;
    static int excelPid;
    static bool NoWhole;

    // Windows commit in use, as a percent of the commit limit.
    static double Commit() { var m = new MS { len = (uint)Marshal.SizeOf<MS>() }; GlobalMemoryStatusEx(ref m); return 100.0 * (m.tf - m.af) / m.tf; }
    static long Priv(int id) { try { using var p = Process.GetProcessById(id); return p.PrivateMemorySize64; } catch { return 0; } }
    static long SelfPriv() { using var p = Process.GetCurrentProcess(); return p.PrivateMemorySize64; }
    static string Mem() => $"excel={Priv(excelPid) / 1048576}MB self={SelfPriv() / 1048576}MB commit={Commit():F1}%";
    static string Col(int n) { var s = ""; while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; } return s; }
    static string Addr(int r1, int c1, int r2, int c2) => $"{Col(c1)}{r1}:{Col(c2)}{r2}";

    static void Monitor()
    {
        var t = new Thread(() =>
        {
            while (true)
            {
                Thread.Sleep(500);
                peakExcel = Math.Max(peakExcel, Priv(excelPid)); peakSelf = Math.Max(peakSelf, SelfPriv());
                if (Commit() >= 90) { if (!stop) Console.WriteLine("STOP: commit >= 90%"); stop = true; }
            }
        }) { IsBackground = true };
        t.Start();
    }

    // Wait at 85%, up to 60 s. Returns false when the step must be skipped.
    static bool Gate(string step)
    {
        for (int i = 0; i < 60 && !stop; i++)
        {
            if (Commit() < 85) return true;
            if (i == 0) Console.WriteLine($"waiting before {step}: commit >= 85%");
            Thread.Sleep(1000);
        }
        Console.WriteLine($"SKIPPED {step}: commit stayed high or stop");
        return false;
    }

    static dynamic StartExcel()
    {
        var sw = Stopwatch.StartNew();
        dynamic xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        xl.Visible = false; xl.DisplayAlerts = false; xl.AutomationSecurity = 3; xl.ScreenUpdating = false; xl.EnableEvents = false;
        GetWindowThreadProcessId(new IntPtr((long)xl.Hwnd), out uint pid); excelPid = (int)pid;
        Console.WriteLine($"excel start {sw.ElapsedMilliseconds}ms own pid={pid}; {Mem()}");
        return xl;
    }

    static void CloseExcel(dynamic xl)
    {
        int id = excelPid;
        try { xl.Quit(); } catch { }
        Marshal.FinalReleaseComObject(xl); xl = null; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        for (int i = 0; i < 15; i++) { if (Priv(id) == 0) { Console.WriteLine("own excel closed"); return; } Thread.Sleep(500); }
        Console.WriteLine($"NOTE own excel {id} still closing; not killed");
    }

    static string Head(string path)
    {
        var b = new byte[16]; using var f = File.OpenRead(path); int n = f.Read(b, 0, 16);
        return Encoding.ASCII.GetString(b, 0, n).Replace('\r', ' ').Replace('\n', ' ');
    }

    static object[,] Synthetic(int startRow, int rows, int cols)
    {
        var a = new object[rows, cols];
        for (int i = 0; i < rows; i++)
        {
            long id = startRow + i; var rng = new Random((int)id);
            for (int c = 0; c < cols; c++)
            {
                int k = c % 10;
                if (k < 5) a[i, c] = Math.Round(rng.NextDouble() * 100000, 3);
                else if (k < 8) a[i, c] = "GROUP-" + rng.Next(1000);
                else a[i, c] = $"ITEM-{id}-{c}";
            }
        }
        return a;
    }

    static void Make(string path, int rows, int cols)
    {
        if (!Gate("make")) return;
        dynamic xl = StartExcel();
        try
        {
            var sw = Stopwatch.StartNew();
            dynamic wb = xl.Workbooks.Add(); dynamic ws = wb.Worksheets[1];
            const int Block = 20000;
            for (int r = 1; r <= rows && !stop; r += Block)
            {
                int n = Math.Min(Block, rows - r + 1);
                ws.Range[Addr(r, 1, r + n - 1, cols)].Value2 = Synthetic(r, n, cols);
            }
            Console.WriteLine($"filled {rows}x{cols} in {sw.ElapsedMilliseconds}ms; {Mem()}");
            sw.Restart(); wb.SaveAs(path, 51); Console.WriteLine($"saved in {sw.ElapsedMilliseconds}ms");
            wb.Close(false);
            Console.WriteLine($"file={new FileInfo(path).Length / 1048576.0:F1}MB first-bytes='{Head(path)}'");
        }
        finally { CloseExcel(xl); }
    }

    static void Stats(string label, List<double> t)
    {
        if (t.Count == 0) return;
        t.Sort();
        Console.WriteLine($"{label}: n={t.Count} median={t[t.Count / 2]:F2}ms p95={t[(int)(t.Count * 0.95)]:F2}ms max={t[^1]:F2}ms");
    }

    static void Bench(string path, bool cache, string cacheDir)
    {
        if (!Gate("open")) return;
        Console.WriteLine($"file={new FileInfo(path).Length / 1048576.0:F2}MB first-bytes='{Head(path)}'");
        dynamic xl = StartExcel();
        string data = null, idx = null;
        try
        {
            var sw = Stopwatch.StartNew();
            dynamic wb = xl.Workbooks.Open(path, 0, true);
            Console.WriteLine($"open read-only {sw.ElapsedMilliseconds}ms; {Mem()}");
            dynamic ws = null; long most = -1; foreach (dynamic sh in wb.Worksheets) { long n = (long)sh.UsedRange.Rows.Count * sh.UsedRange.Columns.Count; if (n > most) { most = n; ws = sh; } } dynamic u = ws.UsedRange;
            int rows = u.Rows.Count, cols = u.Columns.Count;
            Console.WriteLine($"used range {rows}x{cols} cells={(long)rows * cols}");
            var rng = new Random(7);

            // D: screen only, random positions. This is the scrolling case.
            if (Gate("screen"))
            {
                int vr = Math.Min(40, rows), vc = Math.Min(30, cols); var t = new List<double>();
                for (int i = 0; i < 100 && !stop; i++)
                {
                    int r = rng.Next(1, Math.Max(2, rows - vr + 1));
                    sw.Restart(); var a = ws.Range[Addr(r, 1, r + vr - 1, vc)].Value2; t.Add(sw.Elapsed.TotalMilliseconds);
                }
                Stats($"D screen {vr}x{vc}", t); Console.WriteLine($"   after screen reads: {Mem()}");
            }

            // B: block by block, nothing kept.
            foreach (int bs in new[] { 1000, 10000 })
            {
                if (bs > rows || !Gate($"blocks {bs}")) continue;
                peakSelf = 0; sw.Restart(); int blocks = 0; var tb = new List<double>(); var s2 = new Stopwatch();
                for (int r = 1; r <= rows && !stop; r += bs)
                {
                    s2.Restart(); var a = ws.Range[Addr(r, 1, Math.Min(rows, r + bs - 1), cols)].Value2; tb.Add(s2.Elapsed.TotalMilliseconds); blocks++;
                }
                double sec = sw.Elapsed.TotalSeconds;
                Stats($"B block {bs} rows", tb);
                Console.WriteLine($"   total {sec:F2}s = {(long)rows * cols / sec / 1e6:F2}M cells/s; {blocks} blocks; self peak={peakSelf / 1048576}MB");
                GC.Collect();
            }

            // A: everything in one array.
            if (!NoWhole && Gate("whole"))
            {
                peakSelf = 0; long before = SelfPriv(); sw.Restart();
                object all = u.Value2; double sec = sw.Elapsed.TotalSeconds;
                long after = SelfPriv();
                Console.WriteLine($"A whole array {sec:F2}s = {(long)rows * cols / sec / 1e6:F2}M cells/s; self grew {(after - before) / 1048576}MB (~{(double)(after - before) / ((long)rows * cols):F0} bytes/cell); {Mem()}");
                all = null; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            }

            // Search inside Excel for a value that does not exist: worst case.
            if (Gate("find"))
            {
                sw.Restart(); dynamic f = u.Find("NO-SUCH-VALUE-ZZZ"); Console.WriteLine($"Excel Find full miss {sw.ElapsedMilliseconds}ms");
            }

            // Fast copy, made-up data only.
            if (cache && Gate("cache"))
            {
                data = Path.Combine(cacheDir, "cache.bin"); idx = Path.Combine(cacheDir, "cache.idx");
                sw.Restart(); var off = new long[rows + 1];
                using (var fs = new FileStream(data, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16))
                using (var bw = new BinaryWriter(fs, Encoding.UTF8))
                {
                    const int Bs = 10000;
                    for (int r = 1; r <= rows && !stop; r += Bs)
                    {
                        int e = Math.Min(rows, r + Bs - 1);
                        var a = (object[,])ws.Range[Addr(r, 1, e, cols)].Value2;
                        for (int i = 1; i <= e - r + 1; i++)
                        {
                            bw.Flush(); off[r + i - 2] = fs.Position;
                            for (int c = 1; c <= cols; c++)
                                switch (a[i, c])
                                {
                                    case null: bw.Write((byte)0); break;
                                    case double d: bw.Write((byte)1); bw.Write(d); break;
                                    case string s: bw.Write((byte)2); bw.Write(s); break;
                                    default: bw.Write((byte)2); bw.Write(a[i, c].ToString()); break;
                                }
                        }
                    }
                    bw.Flush(); off[rows] = fs.Position;
                }
                using (var fi = new BinaryWriter(File.Create(idx))) foreach (var o in off) fi.Write(o);
                double wsec = sw.Elapsed.TotalSeconds;
                Console.WriteLine($"cache build {wsec:F2}s size data={new FileInfo(data).Length / 1048576.0:F1}MB index={new FileInfo(idx).Length / 1048576.0:F1}MB");
            }
            wb.Close(false);
        }
        finally { CloseExcel(xl); }
        Console.WriteLine($"peak excel private={peakExcel / 1048576}MB");

        // Reopening the fast copy with no Excel at all.
        if (cache && data != null && File.Exists(data))
        {
            long baseMem = SelfPriv(); var sw = Stopwatch.StartNew();
            byte[] raw = File.ReadAllBytes(idx); var off = new long[raw.Length / 8]; Buffer.BlockCopy(raw, 0, off, 0, raw.Length);
            Console.WriteLine($"cache reopen (index only) {sw.ElapsedMilliseconds}ms");
            int rows = off.Length - 1; var rng = new Random(11); var t = new List<double>();
            using var fs = new FileStream(data, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
            var buf = new byte[1 << 20];
            for (int i = 0; i < 500; i++)
            {
                int r = rng.Next(0, rows - 40); sw.Restart();
                fs.Seek(off[r], SeekOrigin.Begin); int need = (int)(off[r + 40] - off[r]); fs.ReadExactly(buf, 0, need);
                using var br = new BinaryReader(new MemoryStream(buf, 0, need), Encoding.UTF8);
                while (br.BaseStream.Position < need) { switch (br.ReadByte()) { case 1: br.ReadDouble(); break; case 2: br.ReadString(); break; } }
                t.Add(sw.Elapsed.TotalMilliseconds);
            }
            Stats("cache screen 40 rows (warm disk)", t);
            sw.Restart(); fs.Seek(0, SeekOrigin.Begin); long hits = 0;
            using (var br = new BinaryReader(new BufferedStream(fs, 1 << 16), Encoding.UTF8, true))
            {
                long len = fs.Length, pos = 0;
                while (pos < len)
                {
                    byte tag = br.ReadByte(); pos++;
                    if (tag == 1) { br.ReadDouble(); pos += 8; }
                    else if (tag == 2) { string s = br.ReadString(); pos += Encoding.UTF8.GetByteCount(s) + (s.Length < 128 ? 1 : 2); if (s.Contains("NO-SUCH-VALUE-ZZZ")) hits++; }
                }
            }
            Console.WriteLine($"cache full search miss {sw.ElapsedMilliseconds}ms hits={hits}; self={SelfPriv() / 1048576}MB (start {baseMem / 1048576}MB)");
        }
    }

    [STAThread]
    static int Main(string[] a)
    {
        Monitor();
        NoWhole = a.Contains("--nowhole");
        Console.WriteLine($"commit at start {Commit():F1}%");
        if (a[0] == "make") Make(a[1], int.Parse(a[2]), int.Parse(a[3]));
        else { bool c = a.Contains("--cache"); if (c && !Path.GetFileName(a[1]).StartsWith("syn")) { Console.WriteLine("REFUSED: --cache is for made-up files only"); return 2; } Bench(a[1], c, c ? Path.GetDirectoryName(a[1]) : null); }
        Console.WriteLine($"commit at end {Commit():F1}%");
        return 0;
    }
}
