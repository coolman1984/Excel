// Turns what the cache holds, plus the edit overlay and the style sidecar, into what the grid shows.
using System.Collections.Concurrent;
using System.Drawing;
using System.Globalization;
using ExcelViewer.Core;

namespace ExcelApp;

/// <summary>A style-only edit is a Clear edit with this text: it changes the look, not the value.</summary>
public static class EditMarks
{
    public const string StyleOnly = "\u0001style";
    public static bool IsStyleOnly(CellEdit e) => e.Kind == EditKind.Clear && e.Text == StyleOnly;
}

public readonly record struct Shown3(string Text, byte Kind, double Num, int? Color, bool Warn, StyleDelta? Style, bool Edited);

public sealed class Cells
{
    readonly CacheData cache; readonly EditSet edits; readonly object editLock;
    readonly Func<StyleSidecar?> sidecar;
    readonly ConcurrentDictionary<string, NumFmt?> formats = new();
    static readonly StyleDelta FormulaStyle = new(Italic: true, FontColor: 0x555555);

    public Cells(CacheData cache, EditSet edits, object editLock, Func<StyleSidecar?> sidecar)
    { this.cache = cache; this.edits = edits; this.editLock = editLock; this.sidecar = sidecar; }

    public bool HasEdits { get { lock (editLock) return edits.All.Count > 0; } }

    public bool TryEdit(int sheet, int row, int col, out CellEdit edit)
    { lock (editLock) return edits.TryGet(sheet, row, col, out edit!); }

    public StyleDelta? FileStyle(int sheet, int row, int col)
    { var sc = sidecar(); return sc == null ? null : StyleStore.StyleAt(sc, sheet, row, col); }

    /// <summary>Format code of the cell in the file ("General" when none).</summary>
    public string FileFormat(int sheet, int row, int col)
    {
        var d = cache.Sheets[sheet]; int id = d.FormatAt(row, col);
        return id >= 0 && id < cache.Formats.Count ? cache.Formats[id].Code : "General";
    }

    public string CurrentFormat(int sheet, int row, int col)
    {
        if (TryEdit(sheet, row, col, out var e) && e.Style?.NumberFormat != null) return e.Style.NumberFormat;
        return FileFormat(sheet, row, col);
    }

    public static string Hex(int rgb) => "#" + (rgb & 0xFFFFFF).ToString("X6", CultureInfo.InvariantCulture);

    public static Dictionary<string, object>? StyleJson(StyleDelta? s)
    {
        if (s == null) return null;
        var d = new Dictionary<string, object>();
        if (s.Bold != null) d["bold"] = s.Bold.Value;
        if (s.Italic != null) d["italic"] = s.Italic.Value;
        if (s.Underline != null) d["underline"] = s.Underline.Value;
        if (s.FontColor != null) d["fontColor"] = Hex(s.FontColor.Value);
        if (s.Fill != null) d["fill"] = Hex(s.Fill.Value);
        if (s.Align != null) d["align"] = s.Align.Value;
        if (s.NumberFormat != null) d["numberFormat"] = s.NumberFormat;
        if (s.Border != null) d["border"] = s.Border;
        return d.Count == 0 ? null : d;
    }

    NumFmt? Compiled(string code) => formats.GetOrAdd(code, c => { try { return new NumFmt(c, cache.Date1904, cache.Names); } catch (Exception) { return null; } });

    Shown3 FormatNumber(string code, double v, StyleDelta? style, bool edited)
    {
        if (string.Equals(code, "General", StringComparison.OrdinalIgnoreCase))
            return new Shown3(v.ToString("R", CultureInfo.CurrentCulture), CellKind.Number, v, null, false, style, edited);
        var s = Compiled(code)?.Number(v);
        if (s == null) return new Shown3(v.ToString("R", CultureInfo.CurrentCulture) + " ⚠", CellKind.Number, v, null, true, style, edited);
        var col = s.Value.Color; int? rgb = col == null ? null : (col.Value.R << 16) | (col.Value.G << 8) | col.Value.B;
        return new Shown3(s.Value.Text.Trim(), CellKind.Number, v, rgb, false, style, edited);
    }

    /// <summary>One cell as shown. Row and col are sheet-relative real indexes.</summary>
    public Shown3 Render(int sheet, SheetReader rd, int row, int col)
    {
        TryEdit(sheet, row, col, out var edit);
        var fileStyle = FileStyle(sheet, row, col);
        bool has = edit != null;
        StyleDelta? style = has && edit!.Style != null ? (fileStyle ?? new StyleDelta()).Merge(edit.Style) : fileStyle;
        if (has && !EditMarks.IsStyleOnly(edit!))
        {
            switch (edit!.Kind)
            {
                case EditKind.Clear: return new Shown3("", CellKind.Empty, 0, null, false, style, true);
                case EditKind.Text: return new Shown3(edit.Text ?? "", CellKind.Text, 0, null, false, style, true);
                case EditKind.Bool: { bool b = edit.Bool ?? edit.Num != 0; return new Shown3(b ? "TRUE" : "FALSE", CellKind.Bool, b ? 1 : 0, null, false, style, true); }
                case EditKind.Formula: return new Shown3(edit.Formula ?? "", CellKind.Text, 0, null, false, FormulaStyle.Merge(style), true);
                case EditKind.Number:
                    {
                        string? code = edit.Style?.NumberFormat;
                        if (code != null) return FormatNumber(code, edit.Num, style, true);
                        var chunk = ColumnChunk.Build(1, new[] { CellKind.Number }, new[] { edit.Num }, null);
                        var v = rd.Show(chunk, 0, row, col);
                        return new Shown3(v.Text, CellKind.Number, edit.Num, v.Color == null ? null : (v.Color.Value.R << 16) | (v.Color.Value.G << 8) | v.Color.Value.B, v.Warn, style, true);
                    }
            }
        }
        var c = rd.Column(row, col); int i = row % rd.Sheet.BlockRows; byte k = c.KindAt(i);
        double num = k is CellKind.Number or CellKind.Bool or CellKind.Error ? c.Nums[i] : 0;
        if (k == CellKind.Number && edit?.Style?.NumberFormat is string nf) return FormatNumber(nf, num, style, has);
        var cv = rd.Show(c, i, row, col);
        if (k == CellKind.Text && cv.Text.Length == 0) k = CellKind.Empty;
        int? color = cv.Color == null ? null : (cv.Color.Value.R << 16) | (cv.Color.Value.G << 8) | cv.Color.Value.B;
        return new Shown3(cv.Text, k, num, color, cv.Warn, style, has);
    }

    public static int Align(Shown3 s) => s.Style?.Align ?? s.Kind switch { CellKind.Number => 1, CellKind.Bool or CellKind.Error => 2, _ => 0 };

    public object?[] Wire(Shown3 s)
    {
        string? color = s.Color == null ? null : Hex(s.Color.Value);
        object? style = StyleJson(s.Style);
        var cell = new List<object?> { s.Text, (int)s.Kind, Align(s), color, s.Warn ? true : null, style };
        while (cell.Count > 3 && cell[^1] == null) cell.RemoveAt(cell.Count - 1);
        return cell.ToArray();
    }

    /// <summary>What the formula bar shows.</summary>
    public string Raw(int sheet, SheetReader rd, int row, int col)
    {
        if (TryEdit(sheet, row, col, out var e) && !EditMarks.IsStyleOnly(e))
            return e.Kind switch
            {
                EditKind.Clear => "", EditKind.Text => e.Text ?? "", EditKind.Formula => e.Formula ?? "",
                EditKind.Bool => (e.Bool ?? e.Num != 0) ? "TRUE" : "FALSE",
                _ => e.Num.ToString("R", CultureInfo.CurrentCulture),
            };
        var c = rd.Column(row, col); int i = row % rd.Sheet.BlockRows;
        return c.KindAt(i) switch
        {
            CellKind.Number => c.Nums[i].ToString("R", CultureInfo.CurrentCulture),
            CellKind.Text => c.TextAt(i),
            CellKind.Bool => c.Nums[i] != 0 ? "TRUE" : "FALSE",
            CellKind.Error => SheetReader.ErrorText(c.Nums[i]),
            _ => "",
        };
    }

    /// <summary>A CellEdit that reproduces the shown value, for fill. Null when the cell is empty and unedited.</summary>
    public CellEdit? ValueEdit(int sheet, SheetReader rd, int row, int col)
    {
        if (TryEdit(sheet, row, col, out var e) && !EditMarks.IsStyleOnly(e)) return e;
        var c = rd.Column(row, col); int i = row % rd.Sheet.BlockRows;
        return c.KindAt(i) switch
        {
            CellKind.Number => new CellEdit(EditKind.Number, Num: c.Nums[i]),
            CellKind.Text => new CellEdit(EditKind.Text, Text: c.TextAt(i)),
            CellKind.Bool => new CellEdit(EditKind.Bool, Num: c.Nums[i], Bool: c.Nums[i] != 0),
            CellKind.Error => new CellEdit(EditKind.Text, Text: SheetReader.ErrorText(c.Nums[i])),
            _ => new CellEdit(EditKind.Clear),
        };
    }

    public static string Address(int firstRow, int firstCol, int row, int col) => CacheBuilder.Letters(firstCol + col) + (firstRow + row).ToString(CultureInfo.InvariantCulture);
}

/// <summary>The cache plus the edits, as the filter engine reads them. Safe to call from several threads.</summary>
public sealed class SheetSource : ICellSource, IDisposable
{
    readonly CacheData cache; readonly int sheet; readonly Cells cells; readonly SheetInfo d;
    ThreadLocal<SheetReader> readers;
    readonly object gate = new();
    public SheetSource(CacheData cache, int sheet, Cells cells)
    {
        this.cache = cache; this.sheet = sheet; this.cells = cells; d = cache.Sheets[sheet];
        readers = NewReaders();
    }
    ThreadLocal<SheetReader> NewReaders() => new ThreadLocal<SheetReader>(() => new SheetReader(cache, d), true);

    /// <summary>Lets go of the cache file (its readers reopen on the next read), so the file can be replaced.</summary>
    public void Release()
    {
        ThreadLocal<SheetReader> old;
        lock (gate) { old = readers; readers = NewReaders(); }
        foreach (var r in old.Values) r.Dispose(); old.Dispose();
    }
    public int Rows => d.Rows; public int Cols => d.Cols;

    public void Read(int col, int fromRow, int count, string[] text, byte[] kind, double[] num)
        => ReadCore(col, fromRow, count, text, kind, num, false);

    public void ReadSort(int col, int fromRow, int count, string[] text, byte[] kind, double[] num)
        => ReadCore(col, fromRow, count, text, kind, num, true);

    void ReadCore(int col, int fromRow, int count, string[] text, byte[] kind, double[] num, bool raw)
    {
        ThreadLocal<SheetReader> tl; lock (gate) tl = readers;
        var rd = tl.Value!; int done = 0;
        while (done < count)
        {
            int row = fromRow + done, ci = row / d.BlockRows, i0 = row % d.BlockRows;
            int n = Math.Min(count - done, d.BlockRows - i0);
            var c = rd.LoadColumn(ci, col);
            for (int k = 0; k < n; k++)
            {
                int i = i0 + k; byte kd = c.KindAt(i); int o = done + k;
                if (kd == CellKind.Empty) { text[o] = ""; kind[o] = 0; num[o] = 0; continue; }
                text[o] = raw ? (kd == CellKind.Text ? c.TextAt(i) : "") : rd.Show(c, i, row + k, col).Text;
                kind[o] = kd; num[o] = kd == CellKind.Text ? 0 : c.Nums[i];
                if (kd == CellKind.Text && text[o].Length == 0) kind[o] = 0;
            }
            done += n;
        }
        if (cells.HasEdits)
            for (int k = 0; k < count; k++)
                if (cells.TryEdit(sheet, fromRow + k, col, out var e))
                {
                    var s = cells.Render(sheet, rd, fromRow + k, col);
                    text[k] = s.Text; kind[k] = s.Kind; num[k] = s.Num;
                }
    }

    public void Dispose() { Release(); lock (gate) readers.Dispose(); }
}
