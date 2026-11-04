// The fast copy's format and its reader. Reading needs no Excel.
// Layout: magic, source fingerprint, column chunks, then a directory, then the directory's position.
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelViewer.Core;

public static class CellKind { public const byte Empty = 0, Number = 1, Text = 2, Bool = 3, Error = 4; }

/// <summary>One column of one block of rows. Never changes once made.</summary>
public sealed class ColumnChunk
{
    public byte[] Kinds; public byte SolidKind;
    public double[] Nums;
    public byte[] Blob; public int[] Offs; public ushort[] Index; // Index: repeated text is stored once

    public byte KindAt(int r) => Kinds == null ? SolidKind : Kinds[r];
    public ReadOnlySpan<byte> TextBytes(int r) { int e = Index != null ? Index[r] : r; return Blob.AsSpan(Offs[e], Offs[e + 1] - Offs[e]); }
    public string TextAt(int r) => Encoding.UTF8.GetString(TextBytes(r));

    public static ColumnChunk Build(int n, byte[] kinds, double[] nums, string[] text)
    {
        var c = new ColumnChunk { Nums = nums };
        byte first = kinds[0]; bool same = true; for (int i = 1; i < n; i++) if (kinds[i] != first) { same = false; break; }
        if (same) c.SolidKind = first; else c.Kinds = kinds;
        if (text == null) return c;
        var map = new Dictionary<string, ushort>(); var list = new List<string>(); var index = new ushort[n]; bool share = true;
        for (int i = 0; i < n; i++)
        {
            string s = text[i] ?? "";
            if (!map.TryGetValue(s, out ushort e)) { if (list.Count >= Math.Max(1, n / 2)) { share = false; break; } e = (ushort)list.Count; map[s] = e; list.Add(s); }
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

    public void Write(BinaryWriter w)
    {
        w.Write(SolidKind);
        w.Write((byte)((Kinds != null ? 1 : 0) | (Nums != null ? 2 : 0) | (Blob != null ? 4 : 0) | (Index != null ? 8 : 0)));
        if (Kinds != null) { w.Write(Kinds.Length); w.Write(Kinds); }
        if (Nums != null) { w.Write(Nums.Length); w.Write(MemoryMarshal.AsBytes(Nums.AsSpan())); }
        if (Blob != null) { w.Write(Blob.Length); w.Write(Blob); w.Write(Offs.Length); w.Write(MemoryMarshal.AsBytes(Offs.AsSpan())); }
        if (Index != null) { w.Write(Index.Length); w.Write(MemoryMarshal.AsBytes(Index.AsSpan())); }
    }

    public static ColumnChunk Read(byte[] b, int len)
    {
        using var br = new BinaryReader(new MemoryStream(b, 0, len));
        var c = new ColumnChunk { SolidKind = br.ReadByte() }; byte f = br.ReadByte();
        if ((f & 1) != 0) c.Kinds = br.ReadBytes(br.ReadInt32());
        if ((f & 2) != 0) { c.Nums = new double[br.ReadInt32()]; br.BaseStream.ReadExactly(MemoryMarshal.AsBytes(c.Nums.AsSpan())); }
        if ((f & 4) != 0) { c.Blob = br.ReadBytes(br.ReadInt32()); c.Offs = new int[br.ReadInt32()]; br.BaseStream.ReadExactly(MemoryMarshal.AsBytes(c.Offs.AsSpan())); }
        if ((f & 8) != 0) { c.Index = new ushort[br.ReadInt32()]; br.BaseStream.ReadExactly(MemoryMarshal.AsBytes(c.Index.AsSpan())); }
        return c;
    }

    /// <summary>A fingerprint used to prove the copy reads back identical. Never shown.</summary>
    public ulong Mix(ulong h, int n)
    {
        const ulong P = 1099511628211UL;
        for (int r = 0; r < n; r++)
        {
            byte k = KindAt(r); h = (h ^ k) * P;
            if (k == CellKind.Text) foreach (byte x in TextBytes(r)) h = (h ^ x) * P;
            else if (k != CellKind.Empty) h = (h ^ (ulong)BitConverter.DoubleToInt64Bits(Nums[r])) * P;
        }
        return h;
    }
}

public enum FormatStatus : byte { NotChecked = 0, MatchesExcel = 1, CannotCarry = 2 }

public sealed class FormatEntry
{
    public string Code; public FormatStatus Status; public int Samples;
    NumFmt fmt;
    public NumFmt Compiled(CacheData c) => fmt ??= new NumFmt(Code, c.Date1904, c.Names);
}

public sealed class SheetInfo
{
    public string Name; public bool Hidden, HideZeros; public int FirstRow, FirstCol, Rows, Cols, BlockRows; public ulong Hash;
    public List<long[]> Chunks = new(); public float[] ColWidths;
    public int[][] RunStart, RunFmt; // per column: which format applies from which row. -1: format not read
    public long Cells => (long)Rows * Cols;

    public int FormatAt(int row, int col)
    {
        var s = RunStart[col]; int lo = 0, hi = s.Length - 1;
        while (lo < hi) { int mid = (lo + hi + 1) / 2; if (s[mid] <= row) lo = mid; else hi = mid - 1; }
        return RunFmt[col][lo];
    }
}

public sealed class CacheData
{
    public const string Magic = "XVCACHE4"; // change this whenever the layout changes
    public long SourceLength, SourceTicks; public bool Date1904;
    public DateNames Names;
    public List<FormatEntry> Formats = new(); public List<SheetInfo> Sheets = new();
    public string Path;

    public bool IsFreshFor(string source) { var f = new FileInfo(source); return f.Exists && f.Length == SourceLength && f.LastWriteTimeUtc.Ticks == SourceTicks; }

    public void WriteDirectory(BinaryWriter w)
    {
        w.Write(Date1904);
        w.Write(Names != null);
        if (Names != null) foreach (var list in new[] { Names.MonthAbbr, Names.MonthFull, Names.DayAbbr, Names.DayFull }) foreach (var s in list) w.Write(s);
        w.Write(Formats.Count);
        foreach (var f in Formats) { w.Write(f.Code); w.Write((byte)f.Status); w.Write(f.Samples); }
        w.Write(Sheets.Count);
        foreach (var d in Sheets)
        {
            w.Write(d.Name); w.Write(d.Hidden); w.Write(d.HideZeros); w.Write(d.FirstRow); w.Write(d.FirstCol); w.Write(d.Rows); w.Write(d.Cols); w.Write(d.BlockRows); w.Write(d.Hash);
            w.Write(d.Chunks.Count); foreach (var o in d.Chunks) foreach (long x in o) w.Write(x);
            for (int k = 0; k < d.Cols; k++)
            {
                w.Write(d.ColWidths[k]); w.Write(d.RunStart[k].Length);
                for (int i = 0; i < d.RunStart[k].Length; i++) { w.Write(d.RunStart[k][i]); w.Write(d.RunFmt[k][i]); }
            }
        }
    }

    /// <summary>Reads the directory. A damaged copy always gives InvalidDataException, so the caller can read the file again.</summary>
    public static CacheData Load(string path)
    {
        try { return LoadCore(path); }
        catch (Exception e) when (e is EndOfStreamException or ArgumentException or OverflowException or OutOfMemoryException or FormatException or IOException and not FileNotFoundException)
        { throw new InvalidDataException("damaged copy", e); }
    }

    static CacheData LoadCore(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
        using var br = new BinaryReader(fs, Encoding.UTF8);
        if (Encoding.ASCII.GetString(br.ReadBytes(8)) != Magic) throw new InvalidDataException("not a current cache");
        var c = new CacheData { Path = path, SourceLength = br.ReadInt64(), SourceTicks = br.ReadInt64() };
        fs.Seek(-8, SeekOrigin.End); fs.Seek(br.ReadInt64(), SeekOrigin.Begin);
        c.Date1904 = br.ReadBoolean();
        if (br.ReadBoolean())
        {
            string[] Read(int n) { var a = new string[n]; for (int i = 0; i < n; i++) a[i] = br.ReadString(); return a; }
            c.Names = new DateNames { MonthAbbr = Read(12), MonthFull = Read(12), DayAbbr = Read(7), DayFull = Read(7) };
        }
        int nf = br.ReadInt32();
        for (int i = 0; i < nf; i++) c.Formats.Add(new FormatEntry { Code = br.ReadString(), Status = (FormatStatus)br.ReadByte(), Samples = br.ReadInt32() });
        int sheets = br.ReadInt32();
        for (int s = 0; s < sheets; s++)
        {
            var d = new SheetInfo { Name = br.ReadString(), Hidden = br.ReadBoolean(), HideZeros = br.ReadBoolean(), FirstRow = br.ReadInt32(), FirstCol = br.ReadInt32(), Rows = br.ReadInt32(), Cols = br.ReadInt32(), BlockRows = br.ReadInt32(), Hash = br.ReadUInt64() };
            int chunks = br.ReadInt32();
            for (int i = 0; i < chunks; i++) { var o = new long[d.Cols + 1]; for (int k = 0; k <= d.Cols; k++) o[k] = br.ReadInt64(); d.Chunks.Add(o); }
            d.ColWidths = new float[d.Cols]; d.RunStart = new int[d.Cols][]; d.RunFmt = new int[d.Cols][];
            for (int k = 0; k < d.Cols; k++)
            {
                d.ColWidths[k] = br.ReadSingle(); int runs = br.ReadInt32(); d.RunStart[k] = new int[runs]; d.RunFmt[k] = new int[runs];
                for (int i = 0; i < runs; i++) { d.RunStart[k][i] = br.ReadInt32(); d.RunFmt[k][i] = br.ReadInt32(); }
            }
            c.Sheets.Add(d);
        }
        return c;
    }

    /// <summary>The source's size and time, read without loading the whole directory.</summary>
    public static bool HeaderMatches(string cachePath, string source)
    {
        if (!File.Exists(cachePath) || !File.Exists(source)) return false;
        try
        {
            using var br = new BinaryReader(new FileStream(cachePath, FileMode.Open, FileAccess.Read, FileShare.Read));
            if (Encoding.ASCII.GetString(br.ReadBytes(8)) != Magic) return false;
            var f = new FileInfo(source);
            return br.ReadInt64() == f.Length && br.ReadInt64() == f.LastWriteTimeUtc.Ticks;
        }
        catch (Exception e) when (e is IOException or EndOfStreamException) { return false; }
    }
}

public readonly record struct CellView(string Text, byte Kind, System.Drawing.Color? Color, bool Warn);

/// <summary>Reads one sheet from the copy. Keeps only a few thousand rows in memory.</summary>
public sealed class SheetReader : IDisposable
{
    readonly FileStream fs; readonly CacheData cache; public readonly SheetInfo Sheet;
    readonly Dictionary<int, ColumnChunk[]> kept = new(); readonly Queue<int> order = new(); readonly int keepChunks;
    byte[] buf = new byte[1 << 20];

    public SheetReader(CacheData cache, SheetInfo sheet)
    {
        this.cache = cache; Sheet = sheet;
        fs = new FileStream(cache.Path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
        keepChunks = Math.Max(64, 5000 / Math.Max(1, sheet.BlockRows));
    }

    public void Dispose() => fs.Dispose();
    public int RowsInChunk(int chunk) => Math.Min(Sheet.BlockRows, Sheet.Rows - chunk * Sheet.BlockRows);

    public ColumnChunk LoadColumn(int chunk, int col)
    {
        long a = Sheet.Chunks[chunk][col], b = Sheet.Chunks[chunk][col + 1]; int len = (int)(b - a);
        if (buf.Length < len) buf = new byte[len * 2];
        fs.Seek(a, SeekOrigin.Begin); fs.ReadExactly(buf, 0, len);
        return ColumnChunk.Read(buf, len);
    }

    public ColumnChunk Column(int row, int col)
    {
        int ci = row / Sheet.BlockRows;
        if (!kept.TryGetValue(ci, out var cols))
        {
            cols = new ColumnChunk[Sheet.Cols]; kept[ci] = cols; order.Enqueue(ci);
            if (order.Count > keepChunks) kept.Remove(order.Dequeue());
        }
        return cols[col] ??= LoadColumn(ci, col);
    }

    public CellView Cell(int row, int col) => Show(Column(row, col), row % Sheet.BlockRows, row, col);

    public string RawText(int row, int col) { var c = Column(row, col); int i = row % Sheet.BlockRows; return c.KindAt(i) == CellKind.Text ? c.TextAt(i) : null; }

    public CellView Show(ColumnChunk c, int i, int row, int col)
    {
        byte k = c.KindAt(i);
        switch (k)
        {
            case CellKind.Empty: return new CellView("", k, null, false);
            case CellKind.Bool: return new CellView(c.Nums[i] != 0 ? "TRUE" : "FALSE", k, null, false);
            case CellKind.Error: return new CellView(ErrorText(c.Nums[i]), k, null, false);
        }
        if (k == CellKind.Number && c.Nums[i] == 0 && Sheet.HideZeros) return new CellView("", k, null, false); // the sheet's "show zeros" setting is off
        int id = Sheet.FormatAt(row, col);
        var entry = id >= 0 ? cache.Formats[id] : null;
        if (k == CellKind.Text)
        {
            string t = c.TextAt(i);
            if (entry == null || entry.Status == FormatStatus.CannotCarry) return new CellView(t, k, null, false);
            return new CellView(entry.Compiled(cache).Text(t) ?? t, k, null, false);
        }
        // A format that did not match Excel is never guessed: show the plain value and mark it.
        if (entry == null || entry.Status != FormatStatus.MatchesExcel)
            return new CellView(c.Nums[i].ToString("R", CultureInfo.CurrentCulture) + " ⚠", k, null, true);
        var s = entry.Compiled(cache).Number(c.Nums[i]);
        return s == null
            ? new CellView(c.Nums[i].ToString("R", CultureInfo.CurrentCulture) + " ⚠", k, null, true)
            : new CellView(s.Value.Text.Trim(), k, s.Value.Color, false);
    }

    public static string ErrorText(double code) => (long)code switch
    {
        -2146826288 => "#NULL!", -2146826281 => "#DIV/0!", -2146826273 => "#VALUE!", -2146826265 => "#REF!",
        -2146826259 => "#NAME?", -2146826252 => "#NUM!", -2146826246 => "#N/A", _ => "#ERROR",
    };

    /// <summary>Do any formats used on this sheet fail to match Excel?</summary>
    public int ColumnsWithFormatWarnings()
    {
        int n = 0;
        for (int k = 0; k < Sheet.Cols; k++)
            if (Sheet.RunFmt[k].Any(id => id < 0 || cache.Formats[id].Status == FormatStatus.CannotCarry)) n++;
        return n;
    }

    public ulong RecomputeHash()
    {
        ulong h = 14695981039346656037UL;
        for (int ci = 0; ci < Sheet.Chunks.Count; ci++) { int n = RowsInChunk(ci); for (int k = 0; k < Sheet.Cols; k++) h = LoadColumn(ci, k).Mix(h, n); }
        return h;
    }
}
