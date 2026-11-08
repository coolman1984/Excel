using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data;

/// <summary>Atomic column-oriented table storage. An uncommitted part is never opened.</summary>
public static class TableFile
{
    internal sealed class Header { public string Name { get; set; } = ""; public int Rows { get; set; } public List<ColHeader> Columns { get; set; } = new(); public Dictionary<string,string> Meta { get; set; } = new(); }
    internal sealed class ColHeader { public string Name { get; set; } = ""; public string Type { get; set; } = ""; public long Offset { get; set; } public long Length { get; set; } public bool HasNulls { get; set; } public int DictCount { get; set; } }
    static long Align(long n) => checked((n + 7) & ~7L);
    static long Length(Column c) => c switch {
        NumberColumn n => checked((long)n.Rows * (n.Nulls == null ? 8 : 9)),
        BoolColumn b => b.Rows,
        TextColumn t => checked((long)t.Rows * 4 + 4 + t.Dict.Sum(x => 4L + Encoding.UTF8.GetByteCount(x))),
        _ => throw new ArgumentException("Unsupported column implementation.") };
    public static void Write(FileRoot root, string relative, Table table, IReadOnlyDictionary<string,string> meta)
    {
        if (table.Rows < 0 || table.Columns.Any(c => Length(c) > int.MaxValue)) throw new UserError("This table is too large to store.");
        var h = new Header { Name = table.Name, Rows = table.Rows, Meta = new(meta) };
        foreach (var c in table.Columns) h.Columns.Add(new ColHeader { Name = c.Name, Type = c.Type.ToString().ToLowerInvariant(), Length = Length(c), HasNulls = c is NumberColumn n && n.Nulls != null, DictCount = c is TextColumn t ? t.Dict.Length : 0 });
        byte[] json = Array.Empty<byte>();
        // Header contains offsets, whose decimal widths can change its length. Iterate to a fixed point.
        for (int i = 0; i < 32; i++) {
            byte[] next = JsonSerializer.SerializeToUtf8Bytes(h, Json.Options);
            long pos = Align(16L + next.Length);
            bool same = true;
            foreach (var c in h.Columns) { if (c.Offset != pos) same = false; c.Offset = pos; pos = Align(checked(pos + c.Length)); }
            json = next;
            if (same) break;
            if (i == 31) throw new InvalidDataException("Table header did not converge.");
        }
        using (var stream = root.OpenWrite(relative)) using (var w = new BinaryWriter(stream, Encoding.UTF8, true)) {
            w.Write(Encoding.ASCII.GetBytes("XVTABLE1")); w.Write(1); w.Write(json.Length); w.Write(json);
            foreach (var (c, ch) in table.Columns.Zip(h.Columns)) {
                while (stream.Position < ch.Offset) w.Write((byte)0);
                switch (c) {
                    case NumberColumn n:
                        stream.Write(MemoryMarshal.AsBytes(n.Values.AsSpan()));
                        if (n.Nulls != null) foreach (bool b in n.Nulls) w.Write((byte)(b ? 1 : 0));
                        break;
                    case TextColumn t:
                        stream.Write(MemoryMarshal.AsBytes(t.Codes.AsSpan())); w.Write(t.Dict.Length);
                        foreach (string s in t.Dict) { byte[] b = Encoding.UTF8.GetBytes(s); w.Write(b.Length); w.Write(b); }
                        break;
                    case BoolColumn b: stream.Write(MemoryMarshal.AsBytes(b.Values.AsSpan())); break;
                }
            }
        }
        root.Commit(relative);
    }
    public static LazyTable Open(FileRoot root, string relative)
    {
        using var s = root.OpenRead(relative); using var r = new BinaryReader(s, Encoding.UTF8, true);
        if (Encoding.ASCII.GetString(r.ReadBytes(8)) != "XVTABLE1" || r.ReadInt32() != 1) throw new InvalidDataException("Invalid table format.");
        int n = r.ReadInt32(); if (n < 2 || n > 16 * 1024 * 1024 || n > s.Length - s.Position) throw new InvalidDataException("Invalid table header.");
        byte[] headerBytes = r.ReadBytes(n);
        using (var doc = JsonDocument.Parse(headerBytes)) {
            if (!doc.RootElement.TryGetProperty("rows", out var rows) || !rows.TryGetInt32(out int rowCount) || rowCount < 0)
                throw new UserError("This table is too large to open.");
            if (doc.RootElement.TryGetProperty("columns", out var columns)) foreach (var col in columns.EnumerateArray())
                if (!col.GetProperty("length").TryGetInt64(out long length) || length > int.MaxValue)
                    throw new UserError("This table is too large to open.");
        }
        var h = JsonSerializer.Deserialize<Header>(headerBytes, Json.Options) ?? throw new InvalidDataException("Invalid table header.");
        if (h.Rows < 0 || h.Columns.Any(c => c.Offset < 16L + n || c.Offset % 8 != 0 || c.Length < 0 || c.Offset > s.Length || c.Length > s.Length - c.Offset)) throw new InvalidDataException("Invalid table section.");
        return new LazyTable(root, relative, h);
    }
    internal static Column ReadColumn(FileRoot root, string relative, int rows, ColHeader h)
    {
        using var s = root.OpenRead(relative); s.Position = h.Offset; using var r = new BinaryReader(s, Encoding.UTF8, true);
        switch (Enum.Parse<ColType>(h.Type, true)) {
            case ColType.Number: case ColType.Date:
                if (h.Length != (long)rows * (h.HasNulls ? 9 : 8)) throw new InvalidDataException("Invalid number section.");
                var values = new double[rows]; s.ReadExactly(MemoryMarshal.AsBytes(values.AsSpan())); bool[]? nulls = null;
                if (h.HasNulls) { nulls = new bool[rows]; for (int i=0;i<rows;i++) nulls[i] = r.ReadByte() != 0; }
                return new NumberColumn(h.Name, values, nulls, Enum.Parse<ColType>(h.Type, true));
            case ColType.Bool:
                if (h.Length != rows) throw new InvalidDataException("Invalid bool section.");
                var b = new sbyte[rows]; s.ReadExactly(MemoryMarshal.AsBytes(b.AsSpan())); return new BoolColumn(h.Name, b);
            case ColType.Text:
                if (h.Length < (long)rows * 4 + 4) throw new InvalidDataException("Invalid text section.");
                var codes = new int[rows]; s.ReadExactly(MemoryMarshal.AsBytes(codes.AsSpan())); int count = r.ReadInt32();
                if (count != h.DictCount || count < 0 || count > (h.Length - (long)rows * 4 - 4) / 4) throw new InvalidDataException("Invalid dictionary.");
                var dict = new string[count]; long end = h.Offset + h.Length;
                for (int i=0;i<count;i++) { int len = r.ReadInt32(); if (len < 0 || len > end - s.Position) throw new InvalidDataException("Invalid dictionary entry."); dict[i] = Encoding.UTF8.GetString(r.ReadBytes(len)); }
                if (s.Position != end || codes.Any(x => x < -1 || x >= count)) throw new InvalidDataException("Invalid text codes.");
                return new TextColumn(h.Name, codes, dict);
            default: throw new InvalidDataException("Unknown column type.");
        }
    }
}
public sealed class LazyTable
{
    readonly FileRoot root; readonly string path; readonly TableFile.Header header;
    readonly Dictionary<int, Column> cache = new(); readonly object gate = new();
    public IReadOnlyDictionary<string,string> Meta => header.Meta;
    public int Rows => header.Rows;
    public IReadOnlyList<ColumnInfo> Columns => header.Columns.Select(c=>new ColumnInfo(c.Name,Enum.Parse<ColType>(c.Type,true))).ToArray();
    internal LazyTable(FileRoot root, string path, TableFile.Header header) { this.root = root; this.path = path; this.header = header; }
    public Table Load(IEnumerable<string>? columns = null)
    {
        var selected = columns == null ? null : new HashSet<string>(columns, StringComparer.OrdinalIgnoreCase);
        lock (gate) {
            var result = new List<Column>();
            for (int i = 0; i < header.Columns.Count; i++) {
                var col = header.Columns[i]; if (selected != null && !selected.Contains(col.Name)) continue;
                if (!cache.TryGetValue(i, out var value)) cache[i] = value = TableFile.ReadColumn(root, path, header.Rows, col);
                result.Add(value);
            }
            return new Table(header.Name, result);
        }
    }
}
