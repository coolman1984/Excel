// PivotTables in the editor: pivot.fields / pivot.run / pivot.place. The range's first row names the fields; its columns are read
// in blocks through the sheet source (edits, formulas' values and inserted rows included) into a table for the Data engine.
using System.Globalization;
using System.Text.Json;
using ExcelApp.Agent;
using ExcelApp.Data;
using ExcelViewer.Core;

namespace ExcelApp;

public sealed partial class Session
{
    void RegisterPivot(Dictionary<string, Func<JsonElement, Task<object?>>> h)
    {
        h["pivot.fields"] = a => Task.Run<object?>(() => PivotFields(a));
        h["pivot.run"] = a => Task.Run<object?>(() => PivotRun(a));
        h["pivot.place"] = a => Task.Run<object?>(() => PivotPlace(a));
    }

    /// <summary>The source range in logical cells: an Excel address (the agent), the selection, or the whole sheet when one cell is selected.</summary>
    (int Sheet, int R1, int C1, int R2, int C2) PivotSource(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        if (st.Map != null) throw new UserError("Clear the filter and the sorting on this sheet first, then make the PivotTable.");
        int rows = RowsOf(sheet), cols = ColsOf(sheet);
        (int R1, int C1, int R2, int C2) r;
        if (Arg.Str(a, "address") is string addr)
        {
            var (x1, y1, x2, y2) = A1.Parse(addr, d.FirstRow + rows - 1, d.FirstCol + cols - 1);
            r = (x1 - d.FirstRow, y1 - d.FirstCol, x2 - d.FirstRow, y2 - d.FirstCol);
        }
        else if (a.TryGetProperty("range", out var rg)) r = Arg.Rect(rg);
        else r = (0, 0, rows - 1, cols - 1);
        if (r.R1 == r.R2 && r.C1 == r.C2) r = (0, 0, rows - 1, cols - 1);   // one cell: the whole data, as Excel takes the current region
        r = (Math.Max(0, r.R1), Math.Max(0, r.C1), Math.Min(rows - 1, r.R2), Math.Min(cols - 1, r.C2));
        if (r.R2 - r.R1 < 1) throw new UserError("Select the data with its header row and at least one row below it.");
        if (r.C2 - r.C1 + 1 > 200) throw new UserError("A PivotTable source can have up to 200 columns. Select the columns you need.");
        return (sheet, r.R1, r.C1, r.R2, r.C2);
    }

    /// <summary>The range as a table: header names from the first row, numbers where a column is mostly numbers, text otherwise.</summary>
    Table PivotTable((int Sheet, int R1, int C1, int R2, int C2) s)
    {
        var src = sheets[s.Sheet].Source; int n = s.R2 - s.R1; var cols = new List<Column>(); var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var head = new string[1]; var hk = new byte[1]; var hn = new double[1];
        for (int c = s.C1; c <= s.C2; c++)
        {
            src.Read(c, s.R1, 1, head, hk, hn);
            string name = head[0].Trim(); if (name.Length == 0) name = "Column " + CacheBuilder.Letters(cache!.Sheets[s.Sheet].FirstCol + c);
            string unique = name; for (int k = 2; !names.Add(unique); k++) unique = name + " " + k;
            var text = new string[n]; var kind = new byte[n]; var num = new double[n];
            for (int at = 0; at < n; at += 65536) { int m = Math.Min(65536, n - at); var t = new string[m]; var kk = new byte[m]; var nn = new double[m]; src.ReadSort(c, s.R1 + 1 + at, m, t, kk, nn); Array.Copy(t, 0, text, at, m); Array.Copy(kk, 0, kind, at, m); Array.Copy(nn, 0, num, at, m); }
            int numbers = kind.Count(k => k == CellKind.Number), filled = kind.Count(k => k != 0);
            if (filled > 0 && numbers >= filled * 0.8)
            {
                var vals = new double[n]; var nulls = new bool[n];
                for (int i = 0; i < n; i++) { if (kind[i] == CellKind.Number) vals[i] = num[i]; else nulls[i] = true; }
                cols.Add(new NumberColumn(unique, vals, nulls.Any(x => x) ? nulls : null));
            }
            else
            {
                var strs = new string?[n];
                for (int i = 0; i < n; i++) strs[i] = kind[i] == 0 ? null : kind[i] == CellKind.Text ? text[i] : kind[i] == CellKind.Bool ? (num[i] != 0 ? "TRUE" : "FALSE") : num[i].ToString("G15", CultureInfo.CurrentCulture);
                cols.Add(TextColumn.From(unique, strs));
            }
        }
        return new Table("pivot", cols);
    }

    string RangeText((int Sheet, int R1, int C1, int R2, int C2) s) { var d = cache!.Sheets[s.Sheet]; return A1.Format(d.FirstRow + s.R1, d.FirstCol + s.C1, d.FirstRow + s.R2, d.FirstCol + s.C2); }

    object PivotFields(JsonElement a)
    {
        gate.Wait();
        var s = PivotSource(a); var t = PivotTable(s);
        return new { range = RangeText(s), rows = t.Rows, fields = t.Columns.Select(c => new { name = c.Name, type = c is NumberColumn ? "number" : "text" }).ToList() };
    }

    static PivotSpec ReadPivotSpec(JsonElement a)
    {
        var rows = a.TryGetProperty("rows", out var r) && r.ValueKind == JsonValueKind.Array ? r.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => x.Length > 0).ToList() : new();
        string? col = Arg.Str(a, "column"); if (string.IsNullOrWhiteSpace(col)) col = null;
        var values = a.TryGetProperty("values", out var v) && v.ValueKind == JsonValueKind.Array
            ? v.EnumerateArray().Select(x => new PivotValue(Arg.Str(x, "field") ?? "", Arg.Str(x, "fn") ?? "sum")).ToList() : new();
        return new PivotSpec(rows, col, values);
    }

    object PivotRun(JsonElement a)
    {
        gate.Wait();
        var s = PivotSource(a); var p = Pivot.Build(PivotTable(s), ReadPivotSpec(a));
        return new { range = RangeText(s), rowFields = p.RowFields, columnField = p.ColumnField, valueLabels = p.ValueLabels, columnKeys = p.ColumnKeys,
                     rows = p.Rows.Select(r => new { keys = r.Keys, cells = r.Cells, total = r.Total }), columnTotals = p.ColumnTotals, grand = p.Grand, truncated = p.Truncated };
    }

    /// <summary>Writes the pivot as values one column to the right of its source (adding columns or rows at the sheet's end when needed).
    /// Refuses to write over cells that are not empty.</summary>
    object PivotPlace(JsonElement a)
    {
        gate.Wait();
        var s = PivotSource(a); var grid = Pivot.Grid(Pivot.Build(PivotTable(s), ReadPivotSpec(a)));
        int h = grid.Count, w = grid.Max(l => l.Count), top = s.R1, left = s.C2 + 2;
        if ((long)h * w > MaxCells) throw new UserError($"The PivotTable is too large to place ({h:N0} x {w:N0} cells).");
        var d = cache!.Sheets[s.Sheet];
        if (d.FirstCol + left + w - 1 > 16384 || d.FirstRow + top + h - 1 > 1_048_576) throw new UserError("There is no room for the PivotTable to the right of the data.");
        lock (readerLock)
        {
            var rd = Reader(s.Sheet); int rows = RowsOf(s.Sheet), cols = ColsOf(s.Sheet);
            for (int r = 0; r < h; r++) for (int c = 0; c < w; c++)
                {
                    int rr = top + r, cc = left + c; if (rr >= rows || cc >= cols) continue;
                    if (cells!.Render(s.Sheet, rd, rr, cc).Kind != CellKind.Empty)
                        throw new UserError($"The cells where the PivotTable goes ({A1.Format(d.FirstRow + top, d.FirstCol + left, d.FirstRow + top + h - 1, d.FirstCol + left + w - 1)}) are not empty. Clear them first.");
                }
        }
        // room at the sheet's end (each is its own undo step, as in Excel)
        int needCols = left + w - ColsOf(s.Sheet), needRows = top + h - RowsOf(s.Sheet);
        lock (editLock)
        {
            if (needCols > 0) Edits.Structure(new StructOp(s.Sheet, false, ColsOf(s.Sheet), needCols, true), d.Rows, d.Cols);
            if (needRows > 0) Edits.Structure(new StructOp(s.Sheet, true, RowsOf(s.Sheet), needRows, true), d.Rows, d.Cols);
        }
        if (needCols > 0 || needRows > 0) AfterStructure(s.Sheet);
        var list = new List<(int, int, CellEdit?)>();
        for (int r = 0; r < h; r++) for (int c = 0; c < grid[r].Count; c++)
            {
                var v = grid[r][c];
                CellEdit edit = v switch
                {
                    null => new CellEdit(EditKind.Clear),
                    double x => new CellEdit(EditKind.Number, Num: x),
                    string t => new CellEdit(EditKind.Text, Text: t, Style: r == 0 || (r == h - 1 && c == 0) ? new StyleDelta(Bold: true) : null),
                    _ => new CellEdit(EditKind.Text, Text: Convert.ToString(v, CultureInfo.CurrentCulture)),
                };
                if (r == h - 1 && edit.Kind == EditKind.Number) edit = edit with { Style = new StyleDelta(Bold: true) };
                list.Add((top + r, left + c, edit));
            }
        Mutate(s.Sheet, sheets[s.Sheet], list);
        return new { placed = A1.Format(d.FirstRow + top, d.FirstCol + left, d.FirstRow + top + h - 1, d.FirstCol + left + w - 1), structure = needCols > 0 || needRows > 0,
                     sheet = s.Sheet, rows = RowsOf(s.Sheet), cols = ColsOf(s.Sheet), dirty = Dirty, canUndo = Edits.CanUndo };
    }
}
