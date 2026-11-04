// The disk cache. OFF unless --disk-cache-approved is passed. Local fixed disk only.
// Prints only times, sizes and counts. Never prints a cell, a name or a value.
// Format: magic, source fingerprint, column chunks, then a directory at the end.
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

sealed class SheetDir
{
    public string Name; public int Rows, Cols, BlockRows; public List<long[]> Chunks = new(); public ulong Hash;
}

sealed class CacheFile
{
    public const string Magic = "XVCACHE1";
    public long SourceLength, SourceTicks; public List<SheetDir> Sheets = new();

    public static CacheFile ReadDirectory(FileStream fs)
    {
        using var br = new BinaryReader(fs, Encoding.UTF8, true);
        fs.Seek(0, SeekOrigin.Begin);
        if (Encoding.ASCII.GetString(br.ReadBytes(8)) != Magic) throw new InvalidDataException("not a cache file");
        var c = new CacheFile { SourceLength = br.ReadInt64(), SourceTicks = br.ReadInt64() };
        fs.Seek(-8, SeekOrigin.End); fs.Seek(br.ReadInt64(), SeekOrigin.Begin);
        int sheets = br.ReadInt32();
        for (int s = 0; s < sheets; s++)
        {
            var d = new SheetDir { Name = br.ReadString(), Rows = br.ReadInt32(), Cols = br.ReadInt32(), BlockRows = br.ReadInt32(), Hash = br.ReadUInt64() };
            int chunks = br.ReadInt32();
            for (int i = 0; i < chunks; i++) { var o = new long[d.Cols + 1]; for (int k = 0; k <= d.Cols; k++) o[k] = br.ReadInt64(); d.Chunks.Add(o); }
            c.Sheets.Add(d);
        }
        return c;
    }

    public static void WriteColumn(BinaryWriter w, ColumnChunk c)
    {
        w.Write(c.SolidKind);
        byte flags = (byte)((c.Kinds != null ? 1 : 0) | (c.Nums != null ? 2 : 0) | (c.Blob != null ? 4 : 0) | (c.Index != null ? 8 : 0));
        w.Write(flags);
        if (c.Kinds != null) { w.Write(c.Kinds.Length); w.Write(c.Kinds); }
        if (c.Nums != null) { w.Write(c.Nums.Length); w.Write(MemoryMarshal.AsBytes(c.Nums.AsSpan())); }
        if (c.Blob != null) { w.Write(c.Blob.Length); w.Write(c.Blob); w.Write(c.Offs.Length); w.Write(MemoryMarshal.AsBytes(c.Offs.AsSpan())); }
        if (c.Index != null) { w.Write(c.Index.Length); w.Write(MemoryMarshal.AsBytes(c.Index.AsSpan())); }
    }

    public static ColumnChunk ReadColumn(byte[] b, int len)
    {
        using var br = new BinaryReader(new MemoryStream(b, 0, len));
        var c = new ColumnChunk { SolidKind = br.ReadByte() }; byte f = br.ReadByte();
        if ((f & 1) != 0) c.Kinds = br.ReadBytes(br.ReadInt32());
        if ((f & 2) != 0) { int n = br.ReadInt32(); c.Nums = new double[n]; br.BaseStream.ReadExactly(MemoryMarshal.AsBytes(c.Nums.AsSpan())); }
        if ((f & 4) != 0) { c.Blob = br.ReadBytes(br.ReadInt32()); int n = br.ReadInt32(); c.Offs = new int[n]; br.BaseStream.ReadExactly(MemoryMarshal.AsBytes(c.Offs.AsSpan())); }
        if ((f & 8) != 0) { int n = br.ReadInt32(); c.Index = new ushort[n]; br.BaseStream.ReadExactly(MemoryMarshal.AsBytes(c.Index.AsSpan())); }
        return c;
    }

    // Fingerprint of one column chunk, used to prove the copy reads back identical. Never printed.
    public static ulong Mix(ulong h, ColumnChunk c, int n)
    {
        const ulong P = 1099511628211UL;
        for (int r = 0; r < n; r++)
        {
            byte k = c.KindAt(r); h = (h ^ k) * P;
            if (k == 2) foreach (byte x in c.TextAt(r)) h = (h ^ x) * P;
            else if (k == 1 || k == 3 || k == 4) h = (h ^ (ulong)BitConverter.DoubleToInt64Bits(c.Nums[r])) * P;
        }
        return h;
    }
}

static class Build
{
    public static string RefuseReason(string file, string dir, bool approved)
    {
        if (!approved) return "the disk cache switch is OFF (default). Memory only.";
        string full = Path.GetFullPath(dir);
        if (new DriveInfo(Path.GetPathRoot(full)).DriveType != DriveType.Fixed) return "cache must be on a fixed local disk";
        if (full.Contains(@"\GitHub\", StringComparison.OrdinalIgnoreCase)) return "cache may not be inside a repository folder";
        string od = Environment.GetEnvironmentVariable("OneDrive");
        if (!string.IsNullOrEmpty(od) && full.StartsWith(od, StringComparison.OrdinalIgnoreCase)) return "cache may not be inside a synced folder";
        return null;
    }

    public static int Run(string file, string dir, bool approved, bool stopOnlyAt90)
    {
        string why = RefuseReason(file, dir, approved);
        if (why != null) { Console.WriteLine("REFUSED: " + why); return 2; }
        if (Sys.Commit() >= 90) { Console.WriteLine("REFUSED: commit >= 90%"); return 3; }
        if (!stopOnlyAt90)
        {
            for (int i = 0; i < 300 && Sys.Commit() >= 85; i++) { if (i == 0) Console.WriteLine("waiting: commit >= 85%"); Thread.Sleep(1000); }
            if (Sys.Commit() >= 85) { Console.WriteLine("SKIPPED: commit stayed >= 85%"); return 3; }
        }
        Console.WriteLine($"commit at start {Sys.Commit():F1}%; wait at 85%: {(stopOnlyAt90 ? "off, by owner's instruction" : "on")}; stop at 90%: on");

        string final = Path.Combine(dir, "real-test.xvc"), tmp = final + ".part"; var src = new FileInfo(file);
        int xpid = 0; long peakExcel = 0; double peakCommit = 0; bool sampling = true, stop = false; var all = Stopwatch.StartNew(); string phase = "open";
        var sampler = new Thread(() =>
        {
            int tick = 0;
            while (Volatile.Read(ref sampling))
            {
                long e = Sys.Priv(xpid); double c = Sys.Commit(); peakExcel = Math.Max(peakExcel, e); peakCommit = Math.Max(peakCommit, c);
                if (c >= 90 && !stop) { stop = true; Console.WriteLine($"STOP: commit {c:F1}% during {phase}"); }
                if (++tick % 15 == 0) Console.WriteLine($"  ... {all.Elapsed.TotalSeconds:F0}s {phase}: excel={e / 1048576}MB commit={c:F1}%");
                Thread.Sleep(1000);
            }
        }) { IsBackground = true }; sampler.Start();

        dynamic xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        dynamic wb = null; int code = 0;
        try
        {
            xl.Visible = false; xl.DisplayAlerts = false; xl.AutomationSecurity = 3; xl.ScreenUpdating = false; xl.EnableEvents = false;
            Sys.GetWindowThreadProcessId(new IntPtr((long)xl.Hwnd), out uint pid); xpid = (int)pid;
            Console.WriteLine($"excel start {all.ElapsedMilliseconds}ms; excel={Sys.Priv(xpid) / 1048576}MB");
            var sw = Stopwatch.StartNew();
            wb = xl.Workbooks.Open(file, 0, true);
            Console.WriteLine($"OPEN read-only {sw.Elapsed.TotalSeconds:F1}s; excel={Sys.Priv(xpid) / 1048576}MB commit={Sys.Commit():F1}%");
            if (stop) { code = 90; return code; }

            // Sheet sizes first, so a file that will not fit is refused before anything is written.
            phase = "measure"; sw.Restart(); var plan = new List<(dynamic ws, int r0, int c0, int rows, int cols)>(); long cells = 0; int idx = 0;
            foreach (dynamic ws in wb.Worksheets)
            {
                idx++; dynamic u = ws.UsedRange; int r0 = u.Row, c0 = u.Column, rows = u.Rows.Count, cols = u.Columns.Count;
                try
                {   // Trim empty formatted space: last real row and column.
                    dynamic lastR = ws.Cells.Find("*", Type.Missing, -4123, Type.Missing, 1, 2);
                    if (lastR == null) { rows = 0; cols = 0; }
                    else
                    {
                        dynamic lastC = ws.Cells.Find("*", Type.Missing, -4123, Type.Missing, 2, 2);
                        rows = Math.Max(0, (int)lastR.Row - r0 + 1); cols = Math.Max(0, (int)lastC.Column - c0 + 1);
                    }
                }
                catch { }
                Console.WriteLine($"sheet#{idx}: rows={rows} cols={cols}");
                plan.Add((ws, r0, c0, rows, cols)); cells += (long)rows * cols;
            }
            long estimate = cells * 12; long free = new DriveInfo(Path.GetPathRoot(Path.GetFullPath(dir))).AvailableFreeSpace;
            Console.WriteLine($"{plan.Count} sheets, {cells} cells; measured in {sw.Elapsed.TotalSeconds:F1}s; estimated cache {estimate / 1048576}MB; disk free {free / 1048576}MB");
            if (estimate > free - (1L << 30)) { Console.WriteLine("STOP: cache would leave under 1 GB free on this disk"); code = 5; return code; }

            phase = "build"; sw.Restart(); long done = 0;
            using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 20))
            using (var w = new BinaryWriter(fs, Encoding.UTF8))
            {
                w.Write(Encoding.ASCII.GetBytes(CacheFile.Magic)); w.Write(src.Length); w.Write(src.LastWriteTimeUtc.Ticks);
                var dirs = new List<SheetDir>(); var ms = new MemoryStream(); idx = 0;
                foreach (var (ws, r0, c0, rows, cols) in plan)
                {
                    idx++; var d = new SheetDir { Name = (string)ws.Name, Rows = rows, Cols = cols, BlockRows = cols == 0 ? 0 : Math.Clamp(20000 / cols, 50, 2000), Hash = 14695981039346656037UL };
                    dirs.Add(d);
                    string firstCol = Sys.Col(c0), lastCol = Sys.Col(c0 + cols - 1);
                    for (int r = 0; r < rows && !stop; r += d.BlockRows)
                    {
                        double c = Sys.Commit();
                        if (c >= 90) { stop = true; break; }
                        if (c >= 85 && !stopOnlyAt90) { int t = 0; while (Sys.Commit() >= 85 && !stop && t++ < 300) Thread.Sleep(1000); if (Sys.Commit() >= 85) { stop = true; break; } }
                        int n = Math.Min(d.BlockRows, rows - r);
                        object v = ws.Range[$"{firstCol}{r0 + r}:{lastCol}{r0 + r + n - 1}"].Value2;
                        var a = v as object[,];
                        if (a == null) { a = (object[,])Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 }); a[1, 1] = v; }
                        var off = new long[cols + 1];
                        for (int k = 0; k < cols; k++)
                        {
                            var kinds = new byte[n]; double[] nums = null; string[] text = null;
                            for (int i = 0; i < n; i++)
                                switch (a[i + 1, k + 1])
                                {
                                    case null: break;
                                    case double x: kinds[i] = 1; (nums ??= new double[n])[i] = x; break;
                                    case string s: kinds[i] = 2; (text ??= new string[n])[i] = s; break;
                                    case bool b: kinds[i] = 3; (nums ??= new double[n])[i] = b ? 1 : 0; break;
                                    default: kinds[i] = 4; (nums ??= new double[n])[i] = 0; break;
                                }
                            var cc = ColumnarStore.Build(n, kinds, nums, text);
                            d.Hash = CacheFile.Mix(d.Hash, cc, n);
                            w.Flush(); off[k] = fs.Position; CacheFile.WriteColumn(w, cc);
                        }
                        w.Flush(); off[cols] = fs.Position; d.Chunks.Add(off); done += (long)n * cols;
                    }
                    if (stop) break;
                }
                if (!stop)
                {
                    w.Flush(); long dirAt = fs.Position; w.Write(dirs.Count);
                    foreach (var d in dirs) { w.Write(d.Name); w.Write(d.Rows); w.Write(d.Cols); w.Write(d.BlockRows); w.Write(d.Hash); w.Write(d.Chunks.Count); foreach (var o in d.Chunks) foreach (long x in o) w.Write(x); }
                    w.Write(dirAt);
                }
            }
            if (stop) { File.Delete(tmp); Console.WriteLine("partial cache deleted"); code = 90; return code; }
            File.Move(tmp, final, true);
            double buildS = sw.Elapsed.TotalSeconds;
            Console.WriteLine($"BUILD {buildS:F1}s ({done / buildS / 1e6:F2}M cells/s); cache {new FileInfo(final).Length / 1048576.0:F1}MB; self peak {Sys.Self() / 1048576}MB");
        }
        catch (Exception e)
        {
            Console.WriteLine("ERROR: " + e.GetType().Name); code = 1;   // the type only: a message could quote file contents
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
        }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch { }
            wb = null;
            try { xl.Quit(); } catch { }
            Marshal.FinalReleaseComObject(xl); xl = null; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            var cl = Stopwatch.StartNew();
            while (Sys.Priv(xpid) != 0 && cl.Elapsed.TotalSeconds < 60) { GC.Collect(); GC.WaitForPendingFinalizers(); Thread.Sleep(200); }
            Console.WriteLine($"own Excel {(Sys.Priv(xpid) == 0 ? "left" : "STILL CLOSING, not killed")} after {cl.Elapsed.TotalSeconds:F1}s; peak excel={peakExcel / 1048576}MB peak commit={peakCommit:F1}%");
            Volatile.Write(ref sampling, false);
        }
        return code;
    }
}

static class Reopen
{
    const int ScreenRows = 40;

    static string Med(List<double> t) { t.Sort(); return $"median={t[t.Count / 2]:F2}ms p95={t[(int)(t.Count * 0.95)]:F2}ms max={t[^1]:F2}ms"; }

    sealed class Sheet
    {
        readonly FileStream fs; readonly SheetDir d; readonly Dictionary<int, ColumnChunk[]> cache = new(); readonly Queue<int> order = new(); byte[] buf = new byte[1 << 20];
        public Sheet(FileStream fs, SheetDir d) { this.fs = fs; this.d = d; }
        public ColumnChunk Column(int chunk, int k)
        {
            long a = d.Chunks[chunk][k], b = d.Chunks[chunk][k + 1]; int len = (int)(b - a);
            if (buf.Length < len) buf = new byte[len * 2];
            fs.Seek(a, SeekOrigin.Begin); fs.ReadExactly(buf, 0, len);
            return CacheFile.ReadColumn(buf, len);
        }
        ColumnChunk[] Chunk(int i)
        {
            if (cache.TryGetValue(i, out var c)) return c;
            c = new ColumnChunk[d.Cols]; for (int k = 0; k < d.Cols; k++) c[k] = Column(i, k);
            cache[i] = c; order.Enqueue(i); if (order.Count > 32) cache.Remove(order.Dequeue());
            return c;
        }
        public string[] Row(int r) { var cc = Chunk(r / d.BlockRows); int i = r % d.BlockRows; var o = new string[d.Cols]; for (int k = 0; k < d.Cols; k++) o[k] = cc[k].Display(i); return o; }
    }

    public static int Run(string path)
    {
        var all = Stopwatch.StartNew();
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
        var c = CacheFile.ReadDirectory(fs); double dirMs = all.Elapsed.TotalMilliseconds;
        var main = c.Sheets.OrderByDescending(s => (long)s.Rows * s.Cols).First(); var sheet = new Sheet(fs, main);
        long sink = 0; for (int r = 0; r < Math.Min(ScreenRows, main.Rows); r++) sink += sheet.Row(r).Length;
        Console.WriteLine($"REOPEN no Excel: directory {dirMs:F1}ms; FIRST SCREEN at {all.Elapsed.TotalMilliseconds:F1}ms; memory {Sys.Self() / 1048576}MB; largest sheet {main.Rows}x{main.Cols}");

        var rng = new Random(9); var t = new List<double>();
        for (int i = 0; i < 500; i++)
        {
            int r = rng.Next(0, Math.Max(1, main.Rows - ScreenRows)); var s = Stopwatch.StartNew();
            for (int y = r; y < Math.Min(main.Rows, r + ScreenRows); y++) sink += sheet.Row(y).Length;
            t.Add(s.Elapsed.TotalMilliseconds);
        }
        Console.WriteLine($"random screen of {ScreenRows} rows: {Med(t)}");

        // Full search through every sheet, chunk by chunk, keeping almost nothing in memory.
        const string Needle = "NO-SUCH-VALUE-ZZZ-7Q"; var nb = Encoding.UTF8.GetBytes(Needle); long hits = 0, cells = 0; ulong ok = 0, bad = 0;
        var sw = Stopwatch.StartNew(); var raw = new byte[1 << 20];
        foreach (var d in c.Sheets) foreach (var o in d.Chunks) for (int k = 0; k < d.Cols; k++)
                {
                    int len = (int)(o[k + 1] - o[k]); if (raw.Length < len) raw = new byte[len * 2];
                    fs.Seek(o[k], SeekOrigin.Begin); fs.ReadExactly(raw, 0, len);
                    if (raw.AsSpan(0, len).IndexOf(nb) >= 0) hits++;
                }
        Console.WriteLine($"FULL SEARCH exact bytes, every sheet: {sw.Elapsed.TotalMilliseconds:F0}ms hits={hits}; memory {Sys.Self() / 1048576}MB");
        sw.Restart(); long ih = 0;
        foreach (var d in c.Sheets) { var sh = new Sheet(fs, d); for (int ci = 0; ci < d.Chunks.Count; ci++) { int n = Math.Min(d.BlockRows, d.Rows - ci * d.BlockRows); for (int k = 0; k < d.Cols; k++) { var col = sh.Column(ci, k); cells += n; if (col.Blob == null) continue; for (int r = 0; r < n; r++) if (col.KindAt(r) == 2 && Encoding.UTF8.GetString(col.TextAt(r)).Contains(Needle, StringComparison.OrdinalIgnoreCase)) ih++; } } }
        Console.WriteLine($"FULL SEARCH ignoring case, as a person would: {sw.Elapsed.TotalMilliseconds:F0}ms hits={ih}; {cells} cells; memory {Sys.Self() / 1048576}MB");

        // Pick a number column and a text column by looking at the first chunk. Their names are never read.
        int numCol = -1, textCol = -1; { int bestN = -1, bestT = -1, n0 = Math.Min(main.BlockRows, main.Rows); for (int k = 0; k < main.Cols; k++) { var col = sheet.Column(0, k); int a = 0, b = 0; for (int r = 0; r < n0; r++) { byte kd = col.KindAt(r); if (kd == 1) a++; else if (kd == 2) b++; } if (a > bestN) { bestN = a; numCol = k; } if (b > bestT) { bestT = b; textCol = k; } } }
        sw.Restart(); var keys = new double[main.Rows]; var ix = new int[main.Rows];
        for (int ci = 0, p = 0; ci < main.Chunks.Count; ci++) { var col = sheet.Column(ci, numCol); int n = Math.Min(main.BlockRows, main.Rows - ci * main.BlockRows); for (int r = 0; r < n; r++, p++) { ix[p] = p; keys[p] = col.KindAt(r) == 1 ? col.Nums[r] : double.NaN; } }
        double loadS = sw.Elapsed.TotalMilliseconds; Array.Sort(keys, ix);
        Console.WriteLine($"SORT by a number column: {sw.Elapsed.TotalMilliseconds:F0}ms (read column {loadS:F0}ms); memory {Sys.Self() / 1048576}MB");
        sw.Restart(); for (int i = 0; i < Math.Min(ScreenRows, main.Rows); i++) sink += sheet.Row(ix[i]).Length;
        Console.WriteLine($"first screen after that sort: {sw.Elapsed.TotalMilliseconds:F1}ms");
        keys = null;
        sw.Restart(); var strs = new string[main.Rows];
        for (int ci = 0, p = 0; ci < main.Chunks.Count; ci++) { var col = sheet.Column(ci, textCol); int n = Math.Min(main.BlockRows, main.Rows - ci * main.BlockRows); for (int r = 0; r < n; r++, p++) { ix[p] = p; strs[p] = col.KindAt(r) == 2 ? Encoding.UTF8.GetString(col.TextAt(r)) : ""; } }
        Array.Sort(strs, ix, StringComparer.Ordinal);
        Console.WriteLine($"SORT by a text column: {sw.Elapsed.TotalMilliseconds:F0}ms; memory {Sys.Self() / 1048576}MB");
        strs = null; ix = null;

        // Proof the copy reads back identical to what the build saw.
        sw.Restart(); bool same = true;
        foreach (var d in c.Sheets)
        {
            var sh = new Sheet(fs, d); ulong h = 14695981039346656037UL;
            for (int ci = 0; ci < d.Chunks.Count; ci++) { int n = Math.Min(d.BlockRows, d.Rows - ci * d.BlockRows); for (int k = 0; k < d.Cols; k++) h = CacheFile.Mix(h, sh.Column(ci, k), n); }
            if (h != d.Hash) same = false;
        }
        Console.WriteLine($"copy reads back identical: {(same ? "YES" : "NO")} ({sw.Elapsed.TotalSeconds:F1}s)");
        GC.KeepAlive(sink);
        Cold(path);
        return same ? 0 : 6;
    }

    // Reads straight from the disk, skipping the Windows file cache, to see a cold-disk cost.
    static unsafe void Cold(string path)
    {
        const FileOptions NoBuffering = (FileOptions)0x20000000; const int Size = 1 << 20;
        long len = new FileInfo(path).Length; void* mem = NativeMemory.AlignedAlloc(Size, 4096);
        try
        {
            using var h = File.OpenHandle(path, FileMode.Open, FileAccess.Read, FileShare.Read, NoBuffering);
            var span = new Span<byte>(mem, Size); var sw = Stopwatch.StartNew();
            RandomAccess.Read(h, span, Math.Max(0, (len - Size) / 4096 * 4096)); RandomAccess.Read(h, span, 0);
            Console.WriteLine($"cold disk: directory end + first 1 MB {sw.Elapsed.TotalMilliseconds:F1}ms");
            sw.Restart(); long off = 0; int n;
            while (off < len && (n = RandomAccess.Read(h, span, off)) > 0) off += n;
            Console.WriteLine($"cold disk: whole cache {sw.Elapsed.TotalSeconds:F2}s ({len / 1048576.0 / sw.Elapsed.TotalSeconds:F0} MB/s)");
        }
        catch (Exception e) { Console.WriteLine("cold test not available: " + e.GetType().Name); }
        finally { NativeMemory.AlignedFree(mem); }
    }
}
