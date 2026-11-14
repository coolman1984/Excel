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
        public Table Get(TableRef r, IReadOnlyCollection<string>? columns) { Asked.Add(columns); return table; }
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
        Engine.Run(new QuerySpec { Source = new("table", "t") }, rec, CancellationToken.None);
        check("scale: a question that shows every column asks for all", rec.Asked[0] == null, "");

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
        for (int i = 0; i < 10; i++) provider.Get(new TableRef("report", report), new[] { "Sales", "C" + i });
        check("scale: the provider keeps only a few unions", provider.CachedCount <= TableProvider.MaxCached, provider.CachedCount.ToString());
    }
}
