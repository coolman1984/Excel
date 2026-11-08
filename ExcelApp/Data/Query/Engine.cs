using System.Globalization;

namespace ExcelApp.Data.Query;

public sealed record QueryResult(Table Table);

public static class Engine
{
    public static QueryResult Run(QuerySpec spec, ITableProvider provider, CancellationToken token)
    {
        Table table = provider.Get(spec.Source);
        foreach (var join in spec.Joins ?? new()) table = Join.Run(table, provider.Get(join.Right), join, token);
        var rows = Enumerable.Range(0, table.Rows).ToArray();
        if (spec.Where is { } where) { var predicate = Pred.Compile(where, table); rows = rows.Where(r => { token.ThrowIfCancellationRequested(); return predicate(r); }).ToArray(); }
        if (spec.Computed is { Count: > 0 })
        {
            var columns = table.Columns.ToList();
            foreach (var item in spec.Computed)
            {
                if (columns.Any(c => string.Equals(c.Name, item.Name, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("Computed column already exists.");
                var expression = Expr.Parse(item.Expr); var values = new double[table.Rows]; var nulls = new bool[table.Rows];
                foreach (int row in rows) { token.ThrowIfCancellationRequested(); var value = expression.Evaluate(table, row); if (value.HasValue) values[row] = value.Value; else nulls[row] = true; }
                columns.Add(new NumberColumn(item.Name, values, nulls)); table = new Table(table.Name, columns.ToArray());
            }
        }
        if ((spec.GroupBy?.Count ?? 0) > 0 || (spec.Aggregates?.Count ?? 0) > 0)
        {
            var groupCols = (spec.GroupBy ?? new()).Select(n => table[n] ?? throw new ArgumentException("Unknown group column.")).ToArray();
            if (groupCols.Length == 2 && groupCols[0] is TextColumn first && groupCols[1] is TextColumn second && spec.Aggregates is { Count: > 0 } fastAgg && fastAgg.All(a => a.Fn is "sum" or "count") && fastAgg.All(a => a.Col == null || table[a.Col] is NumberColumn))
            {
                table = FastGroup(table, rows, first, second, fastAgg, token); rows = Enumerable.Range(0, table.Rows).ToArray();
                if (spec.Having is { } fastHaving) { var predicate = Pred.Compile(fastHaving, table); rows = rows.Where(predicate).ToArray(); }
                goto grouped;
            }
            var groups = new Dictionary<GroupKey, List<int>>();
            foreach (var row in rows) { token.ThrowIfCancellationRequested(); var key = new GroupKey(groupCols.Select(c => KeyValue(c, row)).ToArray()); if (!groups.TryGetValue(key, out var list)) groups[key] = list = new(); list.Add(row); }
            if (groupCols.Length == 0 && groups.Count == 0) groups[new GroupKey(Array.Empty<object?>())] = new();
            var groupsList = groups.Values.ToArray(); var output = new List<Column>();
            foreach (var c in groupCols) output.Add(Gather(c, groupsList.Select(g => g[0]).ToArray()));
            foreach (var aggregate in spec.Aggregates ?? new())
            {
                Column? col = aggregate.Col == null ? null : table[aggregate.Col] ?? throw new ArgumentException("Unknown aggregate column.");
                var values = new double[groupsList.Length]; var nulls = new bool[groupsList.Length];
                for (int i = 0; i < groupsList.Length; i++)
                {
                    var group = groupsList[i]; var valid = col == null ? group : group.Where(r => !col.IsNull(r)).ToList();
                    switch (aggregate.Fn)
                    {
                        case "count": values[i] = valid.Count; break;
                        case "countDistinct": values[i] = col == null ? valid.Count : valid.Select(r => KeyValue(col, r)).Distinct().Count(); break;
                        case "sum": values[i] = col == null ? throw new ArgumentException("Aggregate requires a column.") : valid.Sum(r => col.Num(r)); break;
                        case "avg": case "min": case "max":
                            if (col == null) throw new ArgumentException("Aggregate requires a column.");
                            if (valid.Count == 0) { nulls[i] = true; break; }
                            values[i] = aggregate.Fn == "avg" ? valid.Average(r => col.Num(r)) : aggregate.Fn == "min" ? valid.Min(r => col.Num(r)) : valid.Max(r => col.Num(r)); break;
                        default: throw new ArgumentException("Unknown aggregate.");
                    }
                }
                output.Add(new NumberColumn(aggregate.As, values, nulls));
            }
            table = new Table(table.Name, output); rows = Enumerable.Range(0, table.Rows).ToArray();
            if (spec.Having is { } having) { var predicate = Pred.Compile(having, table); rows = rows.Where(predicate).ToArray(); }
        }
        grouped:
        if (spec.Sort is { Count: > 0 })
        {
            var sorts = spec.Sort.Select(s => (column: table[s.Col] ?? throw new ArgumentException("Unknown sort column."), s.Desc)).ToArray();
            if (sorts.Length == 1 && sorts[0].column is NumberColumn numeric && numeric.Nulls == null)
            {
                var values = numeric.Values; bool descending = sorts[0].Desc;
                Array.Sort(rows, (a, b) => { int cmp = values[a].CompareTo(values[b]); return cmp == 0 ? a.CompareTo(b) : descending ? -cmp : cmp; });
            }
            else Array.Sort(rows, (a, b) => { foreach (var (c, desc) in sorts) { bool an = c.IsNull(a), bn = c.IsNull(b); if (an != bn) return an ? 1 : -1; if (an) continue; int cmp = c.Type == ColType.Text ? StringComparer.OrdinalIgnoreCase.Compare(c.Text(a), c.Text(b)) : c.Num(a).CompareTo(c.Num(b)); if (cmp != 0) return desc ? -cmp : cmp; } return a.CompareTo(b); });
        }
        if (spec.Limit is < 0) throw new ArgumentException("Limit cannot be negative.");
        if (spec.Limit is int limit && rows.Length > limit) rows = rows.Take(limit).ToArray();
        var selected = spec.Columns is { Count: > 0 } ? spec.Columns.Select(n => table[n] ?? throw new ArgumentException("Unknown projected column.")).ToList() : table.Columns.ToList();
        if (spec.GroupBy is not { Count: > 0 } && spec.Aggregates is not { Count: > 0 }) foreach (var c in table.Columns.Where(c => c.Name is "_v" or "_r")) if (!selected.Contains(c)) selected.Add(c);
        return new QueryResult(new Table(table.Name, selected.Select(c => Gather(c, rows)).ToArray()));
    }
    sealed class FastState
    {
        public readonly int Row; public readonly double[] Values;
        public FastState(int row, int count) { Row = row; Values = new double[count]; }
    }
    static Table FastGroup(Table table, int[] rows, TextColumn a, TextColumn b, List<Aggregate> aggregates, CancellationToken token)
    {
        static int[] Canonical(TextColumn c) { var ids = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase); var result = new int[c.Dict.Length]; for (int i = 0; i < result.Length; i++) { if (!ids.TryGetValue(c.Dict[i], out int id)) ids[c.Dict[i]] = id = i; result[i] = id; } return result; }
        var ca = Canonical(a); var cb = Canonical(b); var cols = aggregates.Select(x => x.Col == null ? null : (NumberColumn)table[x.Col]!).ToArray();
        var groups = new Dictionary<(int, int), FastState>();
        foreach (int row in rows)
        {
            if ((row & 65535) == 0) token.ThrowIfCancellationRequested();
            int ac = a.Codes[row], bc = b.Codes[row]; var key = (ac < 0 ? -1 : ca[ac], bc < 0 ? -1 : cb[bc]);
            if (!groups.TryGetValue(key, out var state)) groups.Add(key, state = new FastState(row, aggregates.Count));
            for (int j = 0; j < aggregates.Count; j++) { var col = cols[j]; if (col != null && col.Nulls != null && col.Nulls[row]) continue; if (aggregates[j].Fn == "count") state.Values[j]++; else if (col != null) state.Values[j] += col.Values[row]; else throw new ArgumentException("Aggregate requires a column."); }
        }
        var states = groups.Values.ToArray(); var sample = states.Select(x => x.Row).ToArray(); var output = new List<Column> { Gather(a, sample), Gather(b, sample) };
        for (int j = 0; j < aggregates.Count; j++) { int index = j; output.Add(new NumberColumn(aggregates[j].As, states.Select(x => x.Values[index]).ToArray(), null)); }
        return new Table(table.Name, output);
    }
    static object? KeyValue(Column c, int row) => c.IsNull(row) ? null : c.Type == ColType.Text ? c.Text(row).ToUpperInvariant() : c.Num(row);
    sealed class GroupKey : IEquatable<GroupKey>
    {
        readonly object?[] parts; public GroupKey(object?[] parts) => this.parts = parts;
        public bool Equals(GroupKey? other) => other != null && parts.SequenceEqual(other.parts);
        public override bool Equals(object? obj) => Equals(obj as GroupKey);
        public override int GetHashCode() { var h = new HashCode(); foreach (var p in parts) h.Add(p); return h.ToHashCode(); }
    }
    internal static Column Gather(Column c, int[] rows)
    {
        if (c is TextColumn text) return new TextColumn(c.Name, rows.Select(r => text.Codes[r]).ToArray(), text.Dict);
        if (c is BoolColumn boolean) return new BoolColumn(c.Name, rows.Select(r => boolean.Values[r]).ToArray());
        return new NumberColumn(c.Name, rows.Select(c.Num).ToArray(), rows.Select(c.IsNull).ToArray(), c.Type);
    }
}
