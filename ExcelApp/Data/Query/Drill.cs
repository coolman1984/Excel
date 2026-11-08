using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data.Query;

public sealed record ResultInfo(string ResultId, int Rows, IReadOnlyList<ColumnInfo> Columns, double Ms, bool Groups);
public sealed record LineageCell(string File, string Sheet, string Cell, int Version, DateTime LoadedAt, string Value);
public static class Drill
{
    public static ResultInfo Run(QuerySpec spec, DataServices s, string kind = "query")
    {
        var watch = System.Diagnostics.Stopwatch.StartNew(); var table = Engine.Run(spec, s.Tables, CancellationToken.None).Table; watch.Stop();
        bool grouped = (spec.GroupBy?.Count ?? 0) > 0 || (spec.Aggregates?.Count ?? 0) > 0;
        string id = s.Results.Add(table, new ResultMeta(kind, spec, grouped, spec.Source));
        return new ResultInfo(id, table.Rows, table.Columns.Where(c => !c.Name.StartsWith('_') || (kind == "diff" && c.Name == "_status")).Select(c => new ColumnInfo(c.Name, c.Type)).ToArray(), watch.Elapsed.TotalMilliseconds, grouped);
    }
    public static ResultInfo Open(DataServices s, string id, int row)
    {
        if (!s.Results.TryGet(id, out var table, out var meta) || meta.Spec is not QuerySpec spec || !meta.Grouped || row < 0 || row >= table.Rows) throw new UserError("That grouped result is no longer available.");
        var predicates = new List<object>(); if (spec.Where is { } where) predicates.Add(where);
        foreach (string name in spec.GroupBy ?? new())
        {
            var c = table[name] ?? throw new UserError("A group column is missing.");
            if (c.IsNull(row)) predicates.Add(new { op = "isNull", col = name });
            else predicates.Add(new { op = "eq", col = name, value = c.Type == ColType.Text ? (object)c.Text(row) : c.Num(row) });
        }
        var clone = JsonSerializer.Deserialize<QuerySpec>(JsonSerializer.Serialize(spec, Json.Options), Json.Options)!;
        clone.Where = JsonSerializer.SerializeToElement(new { op = "and", items = predicates }, Json.Options); clone.GroupBy = null; clone.Aggregates = null; clone.Having = null; clone.Sort = null; clone.Limit = null; clone.Columns = null;
        return Run(clone, s, "drill");
    }
    public static LineageCell Lineage(DataServices s, string id, int row, string col)
    {
        if (!s.Results.TryGet(id, out var table, out var meta) || meta.Grouped || row < 0 || row >= table.Rows) throw new UserError("That row is no longer available.");
        var value = table[col] ?? throw new UserError("That column does not exist."); var v = table["_v"]; var r = table["_r"];
        if (v == null || r == null || v.IsNull(row) || r.IsNull(row)) throw new UserError("Source information is unavailable for this cell.");
        var source = s.Tables.SourceOf((int)v.Num(row)) ?? throw new UserError("Source information is unavailable for this cell.");
        if (!source.SourceColumn.TryGetValue(col, out int column) || column < 1) throw new UserError("Source information is unavailable for this column.");
        var letters = ""; while (column > 0) { column--; letters = (char)('A' + column % 26) + letters; column /= 26; }
        return new LineageCell(source.SourcePath, source.Sheet, source.Sheet + "!" + letters + ((int)r.Num(row)).ToString(System.Globalization.CultureInfo.InvariantCulture), source.PeriodVersion, source.LoadedAt, value.Text(row));
    }
}
