// Memory at scale: a question loads only the columns it uses, report unions keep text dictionary-coded,
// and the provider does not keep every union it ever built.
using System.Text.Json;
using ExcelApp.Data;
using ExcelApp.Data.Query;

public static class ScaleTests
{
    sealed class Recording : ITableProvider
    {
        readonly Table table; public readonly List<IReadOnlyCollection<string>?> Asked = new();
        public Recording(Table table) => this.table = table;
        public Table Get(TableRef r) => Get(r, null);
        public Table Get(TableRef r, IReadOnlyCollection<string>? columns) { Asked.Add(columns); return columns == null ? table : new Table(table.Name, table.Columns.Where(c => columns.Contains(c.Name, StringComparer.OrdinalIgnoreCase)).ToArray()); }
        public readonly List<int[]> RowsAsked = new();
        public Table GetRows(TableRef r, IReadOnlyCollection<string>? columns, int[] rows) { RowsAsked.Add(rows); return new Table(table.Name, table.Columns.Select(c => ExcelApp.Data.Query.Engine.Gather(c, rows)).ToArray()); }
        public IReadOnlyList<TableInfo> List() => Array.Empty<TableInfo>();
        public SourceInfo? SourceOf(int id) => null;
        public int SaveAsVersion(string reportTypeId, string period, Table data, string note) => throw new NotSupportedException();
    }

    static HashSet<string> Set(IReadOnlyCollection<string>? c) => c == null ? new() : new(c, StringComparer.OrdinalIgnoreCase);

    public static void Run(Action<string, bool, string> check, bool offline)
    {
        var t = new Table("t", new Column[] {
            new NumberColumn("Sales", new[] { 5d, 20, 30 }, null), TextColumn.From("Region", new[] { "N", "S", "N" }),
            new NumberColumn("Cost", new[] { 1d, 2, 3 }, null), new NumberColumn("Unused", new[] { 0d, 0, 0 }, null) });

        // Engine: which columns does it ask the store for?
        var rec = new Recording(t);
        using var where = JsonDocument.Parse("{\"op\":\"and\",\"items\":[{\"op\":\"gt\",\"col\":\"Sales\",\"value\":1}]}");
        var grouped = Engine.Run(new QuerySpec { Source = new("table", "t"), Where = where.RootElement, Computed = new() { new("Margin", "[Sales] - Cost") },
            GroupBy = new() { "Region" }, Aggregates = new() { new("sum", "Margin", "M") }, Sort = new() { new("M", true) } }, rec, CancellationToken.None).Table;
        var asked = Set(rec.Asked[0]);
        check("scale: a grouped question asks only for its columns", rec.Asked[0] != null && asked.SetEquals(new[] { "Sales", "Region", "Cost", "Margin", "M", "_v", "_r" }), string.Join(",", asked));
        check("scale: the grouped answer is unchanged", grouped.Rows == 2 && grouped["M"]!.Num(0) == 31 && grouped["M"]!.Num(1) == 18, grouped["M"]!.Num(0).ToString());

        rec.Asked.Clear();
        Engine.Run(new QuerySpec { Source = new("table", "t"), Columns = new() { "Region" } }, rec, CancellationToken.None);
        check("scale: a projection asks only for its columns", Set(rec.Asked[0]).SetEquals(new[] { "Region", "_v", "_r" }), string.Join(",", Set(rec.Asked[0])));

        rec.Asked.Clear();
        using var big = JsonDocument.Parse("{\"op\":\"gt\",\"col\":\"Sales\",\"value\":6}");
        var wide = Engine.Run(new QuerySpec { Source = new("table", "t"), Where = big.RootElement, Sort = new() { new("Cost", true) }, Computed = new() { new("Twice", "Sales * 2") }, Limit = 1 }, rec, CancellationToken.None);
        check("scale: SELECT * computes on its own columns first", rec.Asked.Count == 1 && Set(rec.Asked[0]).SetEquals(new[] { "Sales", "Cost", "_v", "_r" }), string.Join(",", Set(rec.Asked[0])));
        check("scale: SELECT * fetches the other columns for the answer rows only", rec.RowsAsked.Count == 1 && rec.RowsAsked[0].SequenceEqual(new[] { 2 }), rec.RowsAsked.Count == 0 ? "none" : string.Join(",", rec.RowsAsked[0]));
        check("scale: SELECT * answer is complete and in order", wide.Table.Rows == 1 && wide.Total == 2 && string.Join(",", wide.Table.Columns.Select(c => c.Name)) == "Sales,Region,Cost,Unused,Twice" && wide.Table["Region"]!.Text(0) == "N" && wide.Table["Twice"]!.Num(0) == 60, string.Join(",", wide.Table.Columns.Select(c => c.Name)));

        rec.Asked.Clear();
        Engine.Run(new QuerySpec { Source = new("table", "t"), Joins = new() { new JoinSpec { Right = new("table", "r"), On = new() { new("Region", "Region") }, Prefix = "r_" } }, Columns = new() { "Sales", "r_Cost" } }, rec, CancellationToken.None);
        check("scale: a join asks both sides for keys and used columns", rec.Asked.Count == 2 && Set(rec.Asked[0]).IsSupersetOf(new[] { "Sales", "Region" }) && Set(rec.Asked[1]).IsSupersetOf(new[] { "Region", "Cost" }) && !Set(rec.Asked[1]).Contains("Unused"), string.Join(",", Set(rec.Asked[1])));

        check("scale: expression columns are found", Set(Expr.Columns("round([Unit price] * Qty, 2) + coalesce(Tax, 0) - 3")).SetEquals(new[] { "Unit price", "Qty", "Tax" }), string.Join(",", Expr.Columns("round([Unit price] * Qty, 2) + coalesce(Tax, 0) - 3")));

        // Store: a report union loads only what was asked, keeps text coded, and the cache stays small
        var root = StoreRoot.ForTest(@"D:\ExcelAppTest\syn-scale-" + Guid.NewGuid().ToString("N"));
        var catalog = new Catalog(root); string report = "scale"; catalog.Update(c => { c.ReportTypes.Add(report); return 0; });
        var provider = new TableProvider(root, catalog);
        provider.SaveAsVersion(report, "2026-01", new Table("p", new Column[] { TextColumn.From("Region", new[] { "N", "S" }), new NumberColumn("Sales", new[] { 1d, 2 }, null), new NumberColumn("Unused", new[] { 9d, 9 }, null) }), "");
        provider.SaveAsVersion(report, "2026-02", new Table("p", new Column[] { TextColumn.From("Region", new[] { "S", "E" }), new NumberColumn("Sales", new[] { 3d, 4 }, null), new NumberColumn("Unused", new[] { 9d, 9 }, null) }), "");
        var some = provider.Get(new TableRef("report", report), new[] { "Region" });
        check("scale: the union loads only asked columns and lineage", some["Unused"] == null && some["Sales"] == null && some["Region"] != null && some["_v"] != null, string.Join(",", some.Columns.Select(c => c.Name)));
        var region = some["Region"] as TextColumn;
        check("scale: union text stays dictionary-coded", region != null && region.Dict.Length == 3 && region.Text(0) == "N" && region.Text(2) == "S" && region.Text(3) == "E", region == null ? "not text" : string.Join(",", region.Dict));
        var all = provider.Get(new TableRef("report", report));
        check("scale: asking for everything still gives everything", all["Unused"] != null && all.Rows == 4 && all["Sales"]!.Num(3) == 4, "");
        // rows fetched across periods, in the answer's order, with a column missing in one period
        provider.SaveAsVersion(report, "2026-03", new Table("p", new Column[] { TextColumn.From("Region", new[] { "W" }), new NumberColumn("Sales", new[] { 5d }, null) }), "");
        provider.Invalidate(report);
        var picked = provider.GetRows(new TableRef("report", report), null, new[] { 4, 0, 3 });
        check("scale: rows fetched across periods keep the asked order", picked.Rows == 3 && picked["Region"]!.Text(0) == "W" && picked["Region"]!.Text(1) == "N" && picked["Sales"]!.Num(2) == 4, picked.Rows + " " + (picked["Region"]?.Text(0) ?? "?"));
        check("scale: a column missing in a period reads as empty", picked["Unused"] != null && picked["Unused"]!.IsNull(0) && picked["Unused"]!.Num(1) == 9, "");
        var lazy = TableFile.Open(root, "reports/scale/2026-01/v1.xvt");
        var direct = lazy.LoadRows(new[] { "Region", "Sales" }, new[] { 1, 0 });
        check("scale: one file reads single rows", direct["Region"]!.Text(0) == "S" && direct["Sales"]!.Num(1) == 1 && direct["Unused"] == null, "");
        var nullable = new Table("n", new Column[] { new NumberColumn("D", new[] { 1d, 2, 3 }, new[] { false, true, false }, ColType.Date), new BoolColumn("B", new sbyte[] { 1, -1, 0 }), TextColumn.From("T", new[] { "a", null, "c" }) });
        TableFile.Write(root, "n.xvt", nullable, new Dictionary<string, string>());
        var nrows = TableFile.Open(root, "n.xvt").LoadRows(null, new[] { 2, 1 });
        check("scale: single-row reads keep nulls, dates, bools and text", nrows["D"]!.Type == ColType.Date && nrows["D"]!.Num(0) == 3 && nrows["D"]!.IsNull(1) && nrows["B"]!.Num(0) == 0 && nrows["B"]!.IsNull(1) && nrows["T"]!.Text(0) == "c" && nrows["T"]!.IsNull(1), "");
        // another process (the app) loads a new version: a long-running reader (the AI server) sees it without being told
        var reader = new TableProvider(root, new Catalog(root));
        int before = reader.Get(new TableRef("report", report), new[] { "Sales" }).Rows;
        new TableProvider(root, new Catalog(root)).SaveAsVersion(report, "2026-04", new Table("p", new Column[] { new NumberColumn("Sales", new[] { 6d, 7 }, null) }), "");
        check("scale: a new version from another process is seen at once", reader.Get(new TableRef("report", report), new[] { "Sales" }).Rows == before + 2, before.ToString());
        for (int i = 0; i < 10; i++) provider.Get(new TableRef("report", report), new[] { "Sales", "C" + i });
        check("scale: the provider keeps only a few unions", provider.CachedCount <= TableProvider.MaxCached, provider.CachedCount.ToString());
    }
}
