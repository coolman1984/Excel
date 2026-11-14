// PivotTables: the data of a range summarised by row fields, an optional column field, and values (sum, count, average,
// min, max), with row, column and grand totals. The grouping runs on the Data mode's engine, so a million rows stay fast.
using System.Globalization;
using ExcelApp.Data;
using ExcelApp.Data.Query;

namespace ExcelApp;

public sealed record PivotValue(string Field, string Fn);
public sealed record PivotSpec(List<string> Rows, string? Column, List<PivotValue> Values);
public sealed record PivotRow(string[] Keys, double?[][] Cells, double?[] Total);
/// <summary>Cells[i][j]: value j at column key i. Totals per value. Column keys are capped (Truncated says so).</summary>
public sealed record PivotResult(string[] RowFields, string? ColumnField, string[] ValueLabels, string[] ColumnKeys, List<PivotRow> Rows, double?[][] ColumnTotals, double?[] Grand, bool Truncated);

public static class Pivot
{
    public const int MaxColumnKeys = 60, MaxRows = 10_000;
    static readonly string[] Fns = { "sum", "count", "avg", "min", "max" };
    public static string Label(PivotValue v) => (v.Fn switch { "sum" => "Sum of ", "count" => "Count of ", "avg" => "Average of ", "min" => "Min of ", _ => "Max of " }) + v.Field;

    sealed class OneTable(Table t) : ITableProvider
    {
        public Table Get(TableRef r) => t;
        public IReadOnlyList<TableInfo> List() => Array.Empty<TableInfo>();
        public SourceInfo? SourceOf(int id) => null;
        public int SaveAsVersion(string a, string b, Table d, string n) => throw new NotSupportedException();
    }

    public static PivotResult Build(Table data, PivotSpec spec, CancellationToken token = default)
    {
        string Field(string name) => data[name]?.Name ?? throw new UserError($"There is no field named \"{name}\". Fields: {string.Join(", ", data.Columns.Select(c => c.Name))}.");
        if (spec.Rows.Count == 0 && spec.Column == null) throw new UserError("Put at least one field in Rows or Columns.");
        if (spec.Values.Count == 0) throw new UserError("Put at least one field in Values.");
        if (spec.Rows.Count > 3) throw new UserError("Use up to three row fields.");
        var rows = spec.Rows.Select(Field).ToList(); string? col = spec.Column == null ? null : Field(spec.Column);
        if (col != null && rows.Contains(col, StringComparer.OrdinalIgnoreCase)) throw new UserError("A field cannot be in Rows and Columns at once.");
        var values = spec.Values.Select(v =>
        {
            if (!Fns.Contains(v.Fn)) throw new UserError("Summarise by sum, count, average, min or max.");
            string f = Field(v.Field);
            if (v.Fn != "count" && data[f] is not NumberColumn) throw new UserError($"\"{f}\" holds text; it can only be counted.");
            return v with { Field = f };
        }).ToList();
        var provider = new OneTable(data);
        Table Run(List<string> group)
        {
            var aggs = values.Select((v, i) => new Aggregate(v.Fn == "count" ? "count" : v.Fn, v.Field, "_v" + i)).ToList();
            return Engine.Run(new QuerySpec { Source = new("table", "pivot"), GroupBy = group.Count == 0 ? null : group, Aggregates = aggs }, provider, token).Table;
        }
        string Key(Table t, string name, int r) { var c = t[name]!; return c.IsNull(r) ? "(blank)" : c.Type == ColType.Text ? c.Text(r) : c.Num(r).ToString("G15", CultureInfo.InvariantCulture); }
        double?[] Vals(Table t, int r) => values.Select((_, i) => { var c = t["_v" + i]!; return c.IsNull(r) ? (double?)null : c.Num(r); }).ToArray();
        int Compare(string a, string b)
        {
            if (a == "(blank)" || b == "(blank)") return a == b ? 0 : a == "(blank)" ? 1 : -1;
            bool na = double.TryParse(a, NumberStyles.Float, CultureInfo.InvariantCulture, out double x), nb = double.TryParse(b, NumberStyles.Float, CultureInfo.InvariantCulture, out double y);
            return na && nb ? x.CompareTo(y) : na ? -1 : nb ? 1 : StringComparer.OrdinalIgnoreCase.Compare(a, b);
        }
        var cmp = Comparer<string>.Create(Compare);

        // column keys
        var colKeys = new List<string>(); bool truncated = false; Table? colTotals = null;
        if (col != null)
        {
            colTotals = Run(new() { col });
            colKeys = Enumerable.Range(0, colTotals.Rows).Select(r => Key(colTotals, col, r)).OrderBy(k => k, cmp).ToList();
            if (colKeys.Count > MaxColumnKeys) { colKeys = colKeys.Take(MaxColumnKeys).ToList(); truncated = true; }
        }
        var colIndex = colKeys.Select((k, i) => (k, i)).ToDictionary(x => x.k, x => x.i, StringComparer.OrdinalIgnoreCase);   // groups ignore case, as Excel does
        int nv = values.Count, nc = Math.Max(1, colKeys.Count);

        var result = new List<PivotRow>();
        if (rows.Count > 0)
        {
            var totals = Run(rows);
            var byKey = new Dictionary<string, PivotRow>(StringComparer.OrdinalIgnoreCase);
            foreach (int r in Enumerable.Range(0, totals.Rows))
            {
                var keys = rows.Select(f => Key(totals, f, r)).ToArray();
                var row = new PivotRow(keys, Enumerable.Range(0, nc).Select(_ => new double?[nv]).ToArray(), Vals(totals, r));
                if (col == null) row.Cells[0] = row.Total;
                byKey[string.Join("\u0001", keys)] = row; result.Add(row);
            }
            if (col != null)
            {
                var cells = Run(rows.Append(col).ToList());
                for (int r = 0; r < cells.Rows; r++)
                {
                    if (!colIndex.TryGetValue(Key(cells, col, r), out int ci)) continue;
                    if (byKey.TryGetValue(string.Join("\u0001", rows.Select(f => Key(cells, f, r))), out var row)) row.Cells[ci] = Vals(cells, r);
                }
            }
            result.Sort((a, b) => { for (int i = 0; i < a.Keys.Length; i++) { int c = Compare(a.Keys[i], b.Keys[i]); if (c != 0) return c; } return 0; });
            if (result.Count > MaxRows) { result = result.Take(MaxRows).ToList(); truncated = true; }
        }
        var colTot = new double?[nc][];
        for (int i = 0; i < nc; i++) colTot[i] = new double?[nv];
        if (colTotals != null) for (int r = 0; r < colTotals.Rows; r++) if (colIndex.TryGetValue(Key(colTotals, col!, r), out int ci)) colTot[ci] = Vals(colTotals, r);
        var grand = Run(new());
        var g = grand.Rows > 0 ? Vals(grand, 0) : new double?[nv];
        if (col == null) colTot[0] = g;
        return new PivotResult(rows.ToArray(), col, values.Select(Label).ToArray(), colKeys.ToArray(), result, colTot, g, truncated);
    }

    /// <summary>The pivot as rows of cells for a sheet: headers, one line per row key, and a Grand Total line (and column).</summary>
    public static List<List<object?>> Grid(PivotResult p)
    {
        var lines = new List<List<object?>>();
        int nv = p.ValueLabels.Length; bool cols = p.ColumnField != null;
        var head = new List<object?>(p.RowFields.Length == 0 ? new object?[] { "" } : p.RowFields.Cast<object?>());
        if (cols) { foreach (var k in p.ColumnKeys) foreach (var v in p.ValueLabels) head.Add(nv == 1 ? k : k + " - " + v); foreach (var v in p.ValueLabels) head.Add(nv == 1 ? "Grand Total" : "Total " + v); }
        else foreach (var v in p.ValueLabels) head.Add(v);
        lines.Add(head);
        foreach (var r in p.Rows)
        {
            var line = new List<object?>(r.Keys.Cast<object?>());
            if (cols) { foreach (var c in r.Cells) foreach (var v in c) line.Add(v); foreach (var v in r.Total) line.Add(v); }
            else foreach (var v in r.Total) line.Add(v);
            lines.Add(line);
        }
        var total = new List<object?> { "Grand Total" };
        for (int i = 1; i < Math.Max(1, p.RowFields.Length); i++) total.Add(null);
        if (cols) { foreach (var c in p.ColumnTotals) foreach (var v in c) total.Add(v); foreach (var v in p.Grand) total.Add(v); }
        else foreach (var v in p.Grand) total.Add(v);
        lines.Add(total);
        return lines;
    }
}
