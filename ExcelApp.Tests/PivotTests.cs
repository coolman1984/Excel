// PivotTables on the Data engine, without a window: groups, a column field, several values, totals, and the sheet layout.
using ExcelApp;
using ExcelApp.Data;

public static class PivotTests
{
    public static void Run(Action<string, bool, string> check)
    {
        var t = new Table("p", new Column[] {
            TextColumn.From("Region", new[] { "North", "South", "North", "East", "south", null }),
            TextColumn.From("Product", new[] { "Pen", "Pen", "Ink", "Ink", "Ink", "Pen" }),
            new NumberColumn("Sales", new[] { 10d, 20, 30, 40, 50, 60 }, null),
            new NumberColumn("Qty", new[] { 1d, 2, 3, 0, 5, 6 }, new[] { false, false, false, true, false, false }) });
        var p = Pivot.Build(t, new PivotSpec(new() { "Region" }, "Product", new() { new("Sales", "sum"), new("Qty", "count") }));
        check("pivot: row keys sorted, blank last, case folded like Excel", string.Join(",", p.Rows.Select(r => r.Keys[0])) == "East,North,South,(blank)", string.Join(",", p.Rows.Select(r => r.Keys[0])));
        check("pivot: column keys sorted", p.ColumnKeys.SequenceEqual(new[] { "Ink", "Pen" }) && p.ValueLabels.SequenceEqual(new[] { "Sum of Sales", "Count of Qty" }), string.Join(",", p.ColumnKeys));
        var north = p.Rows[1]; var south = p.Rows[2];
        check("pivot: cells per column key and value", north.Cells[0][0] == 30 && north.Cells[1][0] == 10 && north.Total[0] == 40 && south.Cells[0][0] == 50 && south.Cells[1][0] == 20 && south.Total[0] == 70, $"{north.Cells[0][0]} {north.Cells[1][0]} {south.Total[0]}");
        check("pivot: count skips empty values", p.Rows[0].Cells[0][1] == 0 && p.Rows[0].Total[1] == 0 && north.Total[1] == 2, $"{p.Rows[0].Total[1]}");
        check("pivot: an empty pair is empty, not zero", p.Rows[0].Cells[1][0] == null, "");
        check("pivot: column totals and the grand total", p.ColumnTotals[0][0] == 120 && p.ColumnTotals[1][0] == 90 && p.Grand[0] == 210 && p.Grand[1] == 5, $"{p.ColumnTotals[0][0]} {p.Grand[0]}");
        var avg = Pivot.Build(t, new PivotSpec(new() { "Product" }, null, new() { new("Sales", "avg"), new("Sales", "max") }));
        check("pivot: average and max without a column field", avg.Rows[0].Keys[0] == "Ink" && avg.Rows[0].Total[0] == 40 && avg.Rows[1].Total[1] == 60 && avg.Grand[0] == 35, $"{avg.Rows[0].Total[0]} {avg.Grand[0]}");
        var two = Pivot.Build(t, new PivotSpec(new() { "Product", "Region" }, null, new() { new("Sales", "sum") }));
        check("pivot: two row fields", two.Rows.Count == 6 && two.Rows[0].Keys.SequenceEqual(new[] { "Ink", "East" }), string.Join("|", two.Rows.Select(r => string.Join("/", r.Keys))));
        check("pivot: text can be counted", Pivot.Build(t, new PivotSpec(new() { "Product" }, null, new() { new("Region", "count") })).Grand[0] == 5, "");
        string Err(PivotSpec s) { try { Pivot.Build(t, s); return "no error"; } catch (UserError e) { return e.Message; } }
        check("pivot: summing text is refused plainly", Err(new(new() { "Product" }, null, new() { new("Region", "sum") })).Contains("only be counted"), "");
        check("pivot: an unknown field names the fields", Err(new(new() { "Nope" }, null, new() { new("Sales", "sum") })).Contains("Fields: Region"), "");
        check("pivot: values are needed", Err(new(new() { "Product" }, null, new())).Contains("Values"), "");
        var grid = Pivot.Grid(p);
        check("pivot: the sheet layout has headers, rows and a Grand Total line", grid.Count == 6 && grid[0].SequenceEqual(new object?[] { "Region", "Ink - Sum of Sales", "Ink - Count of Qty", "Pen - Sum of Sales", "Pen - Count of Qty", "Total Sum of Sales", "Total Count of Qty" }) && (string?)grid[5][0] == "Grand Total" && (double?)grid[5][5] == 210, string.Join(",", grid[0]));
        var big = new Table("b", new Column[] { TextColumn.From("K", Enumerable.Range(0, 1_000_000).Select(i => "k" + (i % 500)).ToArray()), new NumberColumn("V", Enumerable.Repeat(1d, 1_000_000).ToArray(), null) });
        var sw = System.Diagnostics.Stopwatch.StartNew(); var bp = Pivot.Build(big, new PivotSpec(new() { "K" }, null, new() { new("V", "sum") })); sw.Stop();
        Console.WriteLine($"SPEED pivot 1M rows x 500 groups: {sw.Elapsed.TotalSeconds:F3}s");
        check("pivot: a million rows in under two seconds", bp.Rows.Count == 500 && bp.Grand[0] == 1_000_000 && sw.Elapsed.TotalSeconds < 2, sw.Elapsed.TotalSeconds.ToString("F2"));
    }
}
