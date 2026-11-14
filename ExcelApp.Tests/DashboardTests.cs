// The dashboard designer's tiles on the engine: KPI, chart and table tiles from SQL, old query-spec tiles, and errors per tile.
using ExcelApp.Data;
using ExcelApp.Data.Query;

public static class DashboardTests
{
    sealed class One : ITableProvider
    {
        readonly Table t; public One(Table t) => this.t = t;
        public Table Get(TableRef r) => r.Id == "sales" ? t : throw new ExcelApp.UserError("The table was not found.");
        public IReadOnlyList<TableInfo> List() => new[] { new TableInfo("report", "sales", "sales", "", t.Rows, t.Columns.Select(c => new ColumnInfo(c.Name, c.Type)).ToArray(), DateTime.Now) };
        public SourceInfo? SourceOf(int id) => null;
        public int SaveAsVersion(string a, string b, Table d, string n) => throw new NotSupportedException();
    }

    public static void Run(Action<string, bool, string> check)
    {
        var t = new Table("sales", new Column[] { TextColumn.From("Region", new[] { "North", "South", "North" }), new NumberColumn("Sales", new[] { 10d, 20, 30 }, null) });
        var services = new DataServices { Emit = (_, _) => { }, Store = StoreRoot.ForTest(@"D:\ExcelAppTest\syn-dash-" + Guid.NewGuid().ToString("N")[..8]), Config = ConfigRoot.ForTest(@"D:\ExcelAppTest\syn-dashcfg-" + Guid.NewGuid().ToString("N")[..8]), Tables = new One(t) };
        using var dashboards = new Dashboards(services);
        var d = new Dashboard { Id = "board", Name = "Board", Tiles = new() {
            new Tile { Id = "k", Title = "Total", Kind = "kpi", Sql = "SELECT SUM(Sales) AS Total FROM sales", X = 0, Y = 0, W = 3, H = 2 },
            new Tile { Id = "c", Title = "By region", Kind = "chart", ChartType = "column", Sql = "SELECT Region, SUM(Sales) AS Sales FROM sales GROUP BY Region ORDER BY Region", X = 3, Y = 0, W = 6, H = 4 },
            new Tile { Id = "t", Title = "Rows", Kind = "table", Sql = "SELECT * FROM sales ORDER BY Sales DESC", X = 0, Y = 4, W = 12, H = 3 },
            new Tile { Id = "bad", Title = "Broken", Kind = "kpi", Sql = "SELECT Nope FROM sales" },
            new Tile { Id = "old", Title = "Count", Spec = new QuerySpec { Source = new("report", "sales") }, Measure = "rows", Format = "number" } } };
        dashboards.Save(d);
        var back = dashboards.Get("board");
        check("dashboard: tiles keep their kind, query and place", back.Tiles[1].Kind == "chart" && back.Tiles[1].ChartType == "column" && back.Tiles[1].W == 6 && back.Tiles[2].Y == 4 && back.Tiles[0].Sql!.StartsWith("SELECT SUM"), "");
        var v = dashboards.Refresh("board").ToDictionary(x => x.TileId);
        check("dashboard: a KPI from SQL", v["k"].Value == 60 && v["k"].Error == null, v["k"].Error ?? v["k"].Value?.ToString() ?? "null");
        check("dashboard: a chart tile carries its rows", v["c"].Data != null && v["c"].Data!.Rows.Count == 2 && (string?)v["c"].Data!.Rows[0][0] == "North" && (double?)v["c"].Data!.Rows[0][1] == 40, v["c"].Error ?? "");
        check("dashboard: a table tile carries its rows in order", v["t"].Data != null && v["t"].Data!.Rows.Count == 3 && (double?)v["t"].Data!.Rows[0][1] == 30, "");
        check("dashboard: a broken query is a plain sentence on that tile only", v["bad"].Error != null && v["bad"].Error!.Contains("Nope") && v["k"].Error == null, v["bad"].Error ?? "");
        check("dashboard: older tiles still work", v["old"].Value == 3, v["old"].Error ?? "");
    }
}
