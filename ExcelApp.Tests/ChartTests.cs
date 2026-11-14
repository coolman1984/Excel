// Charts without a window: what a range plots (Excel's rules), undo and redo, and charts following inserted or deleted rows.
using ExcelApp;

public static class ChartTests
{
    static Func<int, int, (byte, double, string)> Grid(object?[,] g) => (r, c) => g[r, c] switch
    {
        null => ((byte)0, 0, ""),
        double d => ((byte)1, d, d.ToString(System.Globalization.CultureInfo.InvariantCulture)),
        int i => ((byte)1, i, i.ToString(System.Globalization.CultureInfo.InvariantCulture)),
        DateTime t => ((byte)1, t.ToOADate(), t.ToString("yyyy-MM-dd")),
        string s => ((byte)2, 0, s),
        _ => ((byte)0, 0, ""),
    };

    public static void Run(Action<string, bool, string> check)
    {
        var sales = new object?[,] { { "Region", "Sales", "Cost" }, { "North", 40, 10 }, { "South", 20, 5 }, { "East", 30, null } };
        var p = ChartData.Read(4, 3, Grid(sales), "column");
        check("chart data: header names the series, first column the categories", p.Categories.SequenceEqual(new[] { "North", "South", "East" }) && p.Series.Length == 2 && p.Series[0].Name == "Sales" && p.Series[1].Name == "Cost", string.Join(",", p.Categories));
        check("chart data: numbers, and an empty cell is a gap", p.Series[0].Values.SequenceEqual(new double?[] { 40, 20, 30 }) && p.Series[1].Values[2] == null, "");
        var topLeftEmpty = new object?[,] { { null, "2025", "2026" }, { "Q1", 5, 7 }, { "Q2", 6, 8 } };
        var p2 = ChartData.Read(3, 3, Grid(topLeftEmpty), "line");
        check("chart data: an empty top-left cell means row and column headers", p2.Categories.SequenceEqual(new[] { "Q1", "Q2" }) && p2.Series[0].Name == "2025", string.Join(",", p2.Categories));
        var numbersOnly = new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 } };
        var p3 = ChartData.Read(2, 2, Grid(numbersOnly), "column");
        check("chart data: numbers only, every column a series, categories numbered", p3.Series.Length == 2 && p3.Categories.SequenceEqual(new[] { "1", "2" }) && p3.Series[0].Name == "Series 1", p3.Series.Length.ToString());
        var dates = new object?[,] { { "Day", "Visits" }, { new DateTime(2026, 1, 1), 5 }, { new DateTime(2026, 1, 2), 9 } };
        var p4 = ChartData.Read(3, 2, Grid(dates), "line");
        check("chart data: dates are categories, shown as Excel shows them", p4.Categories.SequenceEqual(new[] { "2026-01-01", "2026-01-02" }) && p4.Series.Single().Name == "Visits", string.Join(",", p4.Categories));
        var xy = new object?[,] { { "Height", "Weight" }, { 160.0, 55.0 }, { 180.0, 80.0 } };
        var p5 = ChartData.Read(3, 2, Grid(xy), "scatter");
        check("chart data: scatter takes the first number column as x", p5.X != null && p5.X.SequenceEqual(new double?[] { 160, 180 }) && p5.Series.Single().Name == "Weight", "");
        var wide = new object?[2, 12]; for (int c = 0; c < 12; c++) { wide[0, c] = "S" + c; wide[1, c] = (double)c; }
        var p6 = ChartData.Read(2, 12, Grid(wide), "column");
        check("chart data: at most eight series, and it says so", p6.Series.Length == ChartData.MaxSeries && p6.Truncated, p6.Series.Length.ToString());

        // ---- the edit set: charts with undo, and moving with inserted or deleted rows
        var e = new EditSet();
        var spec = new ChartSpec("c1", 0, "column", "Sales", 0, 0, 3, 2, 100, 80, 520, 320);
        e.SetChart(null, spec);
        check("chart edits: added, and unsaved work", e.Chart("c1") == spec && e.Dirty, "");
        e.SetChart(spec, spec with { Type = "line", X = 300 });
        e.Undo();
        check("chart edits: undo a change", e.Chart("c1")!.Type == "column" && e.Chart("c1")!.X == 100 && e.LastWasChart, "");
        e.Redo();
        check("chart edits: redo it", e.Chart("c1")!.Type == "line", "");
        e.Structure(new StructOp(0, true, 2, 2, true), 10, 4);
        check("chart edits: rows inserted inside its range make it longer", e.Chart("c1") is { R1: 0, R2: 5 }, e.Chart("c1")!.R2.ToString());
        e.Structure(new StructOp(0, false, 0, 1, true), 10, 4);
        check("chart edits: a column inserted before it moves it right", e.Chart("c1") is { C1: 1, C2: 3 }, "");
        e.Undo(); e.Undo();
        check("chart edits: undoing the inserts puts the range back", e.Chart("c1") is { R1: 0, R2: 3, C1: 0, C2: 2 }, "");
        e.Structure(new StructOp(0, true, 0, 4, false), 10, 4);
        check("chart edits: deleting all its rows deletes the chart", e.Chart("c1") == null, "");
        e.Undo();
        check("chart edits: undo brings it back", e.Chart("c1") != null, "");
        e.SetChart(e.Chart("c1"), null);
        check("chart edits: delete", e.Chart("c1") == null && e.Charts.Count == 0, "");

        // ---- the journal keeps charts
        e.Undo();
        var snap = EditJournal.Take(Path.Combine(Path.GetTempPath(), "syn-not-there.xlsx"), e);
        var back = System.Text.Json.JsonSerializer.Deserialize<EditJournal.Snapshot>(System.Text.Json.JsonSerializer.Serialize(snap))!;
        var r = new EditSet(); EditJournal.Restore(back, r, s => s == 0 ? (10, 4) : (-1, -1));
        check("chart edits: the journal brings charts back", r.Chart("c1") is { Type: "line", Title: "Sales" }, "");
    }
}
