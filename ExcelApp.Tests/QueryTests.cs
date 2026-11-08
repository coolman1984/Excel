using ExcelApp.Data;
using ExcelApp.Data.Query;
using System.Text.Json;

public static class QueryTests
{
    /// <summary>Offline checks always; checks that start a hidden Excel only when offline is false.</summary>
    public static void Run(Action<string, bool, string> check, bool offline)
    {
        var table = new Table("synthetic", new Column[] {
            new NumberColumn("Amount", new double[] { 10, 0, 2 }, new[] { false, false, true }),
            new NumberColumn("Previous amount", new double[] { 5, 0, 4 }, null)
        });
        double? Eval(string s, int row) => Expr.Parse(s).Evaluate(table, row);
        check("expression arithmetic", Eval("[Amount] + 2 * 3", 0) == 16, "");
        check("expression percent", Eval("pct(Amount, [Previous amount])", 0) == 100, "");
        check("expression divide by zero", Eval("Amount / [Previous amount]", 1) == null, "");
        check("expression null propagation", Eval("Amount + 1", 2) == null, "");
        check("expression coalesce", Eval("coalesce(Amount, 4)", 2) == 4, "");
        check("expression round", Eval("round(-1.25, 1)", 0) == -1.3, "");
        bool invalid = false; try { Expr.Parse("abs(1, 2)"); } catch (FormatException) { invalid = true; }
        check("expression invalid arity", invalid, "");
        var today = new DateTime(2026, 11, 7);
        check("relative previous week", RelativeDate.TryResolve("startOfLastWeek", today, out var date) && date == new DateTime(2026, 10, 26), "");
        check("relative invalid", !RelativeDate.TryResolve("today--1", today, out _), "");
        var predicates = new Table("synthetic", new Column[] {
            TextColumn.From("Name", new string?[] { "Alpha", "alpha", null, "Beta" }),
            new NumberColumn("Value", new double[] { 10, 11, 0, 20 }, new[] { false, false, true, false }),
            new NumberColumn("Other", new double[] { 10, 10, 4, 20 }, null),
            new NumberColumn("Day", new double[] { today.ToOADate(), today.AddDays(-1).ToOADate(), today.ToOADate(), today.ToOADate() }, null, ColType.Date)
        });
        bool TestPred(string json, int row) { using var doc = JsonDocument.Parse(json); return Pred.Compile(doc.RootElement, predicates, today)(row); }
        check("predicate logical and comparison", TestPred("{\"op\":\"and\",\"items\":[{\"op\":\"ge\",\"col\":\"Value\",\"value\":10},{\"op\":\"not\",\"item\":{\"op\":\"blank\",\"col\":\"Name\"}}]}", 0), "");
        check("predicate case insensitive", TestPred("{\"op\":\"contains\",\"col\":\"Name\",\"value\":\"LPH\"}", 0), "");
        check("predicate duplicates", TestPred("{\"op\":\"duplicate\",\"cols\":[\"Name\"]}", 1) && !TestPred("{\"op\":\"duplicate\",\"cols\":[\"Name\"]}", 3), "");
        check("predicate relative date", TestPred("{\"op\":\"eq\",\"col\":\"Day\",\"value\":\"today-1\"}", 1), "");
        check("predicate null excluded", !TestPred("{\"op\":\"ne\",\"col\":\"Value\",\"value\":3}", 2), "");
        check("predicate tolerance", TestPred("{\"op\":\"colCompare\",\"col\":\"Value\",\"other\":\"Other\",\"cmp\":\"eq\",\"tolerancePct\":10}", 1), "");
        check("predicate between", TestPred("{\"op\":\"between\",\"col\":\"Value\",\"a\":10,\"b\":11}", 1), "");
        check("predicate in", TestPred("{\"op\":\"in\",\"col\":\"Name\",\"values\":[\"ALPHA\",\"gamma\"]}", 0), "");
        check("predicate or", TestPred("{\"op\":\"or\",\"items\":[{\"op\":\"blank\",\"col\":\"Name\"},{\"op\":\"lt\",\"col\":\"Value\",\"value\":10}]}", 2), "");
        check("predicate begins ends", TestPred("{\"op\":\"begins\",\"col\":\"Name\",\"value\":\"al\"}", 0) && TestPred("{\"op\":\"ends\",\"col\":\"Name\",\"value\":\"HA\"}", 1), "");
        check("predicate not blank", TestPred("{\"op\":\"notBlank\",\"col\":\"Name\"}", 0) && !TestPred("{\"op\":\"notBlank\",\"col\":\"Name\"}", 2), "");
        var pageTable = new Table("synthetic", new Column[] { new NumberColumn("N", new[] { 1.123456789012, 0d }, new[] { false, true }), new NumberColumn("Date", new[] { today.ToOADate(), today.ToOADate() }, null, ColType.Date), new BoolColumn("Bool", new sbyte[] { 1, -1 }), TextColumn.From("Text", new string?[] { "abc", null }), new NumberColumn("_v", new[] { 1d, 2d }, null) });
        var store = new ResultStore(); string pageId = store.Add(pageTable, new ResultMeta("diff", null, false, null));
        var page = Format.Page(store, pageId, 0, 2);
        check("page formatting and hidden columns", page.Rows[0].SequenceEqual(new[] { "1.123456789", "2026-11-07", "TRUE", "abc" }) && page.Kinds[0].SequenceEqual(new[] { 1, 1, 3, 2 }) && page.Kinds[1][0] == 0 && page.Rows[0].Length == 4, "");
        var diffTable = new Table("synthetic", new Column[] { TextColumn.From("Value", new string?[] { "X" }), TextColumn.From("_status", new string?[] { "changed" }), new NumberColumn("_v", new[] { 1d }, null), new NumberColumn("_r", new[] { 2d }, null) });
        string diffId = store.Add(diffTable, new ResultMeta("diff", null, false, null));
        string queryId = store.Add(diffTable, new ResultMeta("query", null, false, null));
        check("diff page includes only status reserved column", Format.Page(store, diffId, 0, 1).Rows[0].SequenceEqual(new[] { "X", "changed" }) && Format.Page(store, queryId, 0, 1).Rows[0].Length == 1, "");
        check("page bounds", Format.Page(store, pageId, 9, 5).Rows.Length == 0, "");
        var fake = new FakeProvider(predicates);
        var grouped = Engine.Run(new QuerySpec { Source = new("table", "x"), GroupBy = new() { "Name" }, Aggregates = new() { new("sum", "Value", "Total"), new("count", null, "Rows") }, Sort = new() { new("Total", true) } }, fake, CancellationToken.None).Table;
        check("engine group and sort", grouped.Rows == 3 && grouped["Total"]!.Num(0) == 21 && grouped["Rows"]!.Num(0) == 2, "");
        using var filter = JsonDocument.Parse("{\"op\":\"gt\",\"col\":\"Value\",\"value\":10}");
        var filtered = Engine.Run(new QuerySpec { Source = new("table", "x"), Where = filter.RootElement, Computed = new() { new("Double", "Value * 2") }, Sort = new() { new("Value", true) }, Limit = 1, Columns = new() { "Double" } }, fake, CancellationToken.None).Table;
        check("engine filter computed projection", filtered.Rows == 1 && filtered["Double"]!.Num(0) == 40, "");
        const int large = 10_000_000; var codesA = new int[large]; var codesB = new int[large]; var numbers = new double[large];
        for (int i = 0; i < large; i++) { codesA[i] = i % 10; codesB[i] = (i / 10) % 10; numbers[i] = 1; }
        var big = new Table("synthetic", new Column[] { new TextColumn("A", codesA, Enumerable.Range(0, 10).Select(x => x.ToString()).ToArray()), new TextColumn("B", codesB, Enumerable.Range(0, 10).Select(x => x.ToString()).ToArray()), new NumberColumn("N", numbers, null) });
        var bigProvider = new FakeProvider(big);
        var timer = System.Diagnostics.Stopwatch.StartNew();
        var bigResult = Engine.Run(new QuerySpec { Source = new("table", "x"), GroupBy = new() { "A", "B" }, Aggregates = new() { new("sum", "N", "Sum"), new("count", null, "Count") } }, bigProvider, CancellationToken.None).Table;
        timer.Stop(); Console.WriteLine($"SPEED group-by 10M: {timer.Elapsed.TotalSeconds:F3}s (target <1s)");
        check("speed group 10M", bigResult.Rows == 100 && bigResult["Count"]!.Num(0) == 100000 && timer.Elapsed.TotalSeconds < 1, "");
        var million = new Table("synthetic", new Column[] { new NumberColumn("N", Enumerable.Range(0, 1_000_000).Select(i => (double)(i % 1000)).ToArray(), null) });
        using var speedFilter = JsonDocument.Parse("{\"op\":\"ge\",\"col\":\"N\",\"value\":500}");
        timer.Restart(); var sorted = Engine.Run(new QuerySpec { Source = new("table", "x"), Where = speedFilter.RootElement, Sort = new() { new("N", true) } }, new FakeProvider(million), CancellationToken.None).Table; timer.Stop();
        Console.WriteLine($"SPEED filter+sort 1M: {timer.Elapsed.TotalSeconds:F3}s (target <1s)");
        check("speed filter sort 1M", sorted.Rows == 500_000 && sorted["N"]!.Num(0) == 999 && sorted["N"]!.Num(sorted.Rows - 1) == 500 && timer.Elapsed.TotalSeconds < 1, "");
        check("fuzzy normalize digit runs", Fuzzy.Normalize(" ab-00012 / 003 ") == "AB123", "");
        check("fuzzy jaro winkler", Math.Abs(Fuzzy.JaroWinkler("MARTHA", "MARHTA") - 0.9611111111111111) < 1e-9, "");
        var jl = new Table("left", new Column[] { TextColumn.From("Key", new string?[] { " AB-001 ", "MISS", "ALPHA" }) });
        var jr = new Table("right", new Column[] { TextColumn.From("Key", new string?[] { "ab1", "ALPXA" }) });
        var joinSpec = new JoinSpec { On = new() { new("Key", "Key") }, Kind = "left", Fuzzy = "similar", MinScore = .8 };
        var joined = Join.Run(jl, jr, joinSpec, CancellationToken.None);
        check("join normalized similar missing", joined.Rows == 3 && joined["_match"]!.Text(0) == "normalized" && joined["_match"]!.Text(1) == "none" && joined["_match"]!.Text(2) == "similar", "");
        joinSpec.Kind = "anti"; var anti = Join.Run(jl, jr, joinSpec, CancellationToken.None);
        check("join anti", anti.Rows == 1 && anti["Key"]!.Text(0) == "MISS", "");
        var duplicates = new Table("right", new Column[] { TextColumn.From("Key", new string?[] { "A", "A" }) });
        var one = new Table("left", new Column[] { TextColumn.From("Key", new string?[] { "A", null }) });
        var multi = Join.Run(one, duplicates, new JoinSpec { On = new() { new("Key", "Key") }, Kind = "left" }, CancellationToken.None);
        check("join duplicate and null keys", multi.Rows == 3 && multi["_match"]!.Text(0) == "exact" && multi["_match"]!.Text(2) == "none" && multi["right.Key"]!.IsNull(2), "");
        var leftComposite = new Table("l", new Column[] { TextColumn.From("A", new string?[] { "CODE" }), TextColumn.From("B", new string?[] { "ALPHA" }) });
        var rightComposite = new Table("r", new Column[] { TextColumn.From("A", new string?[] { "CODE" }), TextColumn.From("B", new string?[] { "ALPXA" }) });
        var composite = Join.Run(leftComposite, rightComposite, new JoinSpec { On = new() { new("A", "A"), new("B", "B") }, Kind = "inner", Fuzzy = "similar", MinScore = .8 }, CancellationToken.None);
        check("join composite similar", composite.Rows == 1 && composite["_match"]!.Text(0) == "similar", "");
        var leftKeys = new string?[1_000_000]; var rightKeys = new string?[1_000_000];
        for (int i = 0; i < leftKeys.Length; i++) { leftKeys[i] = "a-" + i.ToString("D8"); rightKeys[i] = "A_" + i.ToString(); }
        var leftBig = new Table("left", new Column[] { TextColumn.From("Key", leftKeys) });
        var rightBig = new Table("right", new Column[] { TextColumn.From("Key", rightKeys) });
        timer.Restart(); var joinBig = Join.Run(leftBig, rightBig, new JoinSpec { On = new() { new("Key", "Key") }, Kind = "left", Fuzzy = "normalized" }, CancellationToken.None); timer.Stop();
        Console.WriteLine($"SPEED messy join 1M x 1M: {timer.Elapsed.TotalSeconds:F3}s (target <10s)");
        check("speed messy join 1M", joinBig.Rows == 1_000_000 && joinBig["_match"]!.Text(999999) == "normalized" && timer.Elapsed.TotalSeconds < 10, "");
        var spec = JsonSerializer.Deserialize<QuerySpec>("{\"source\":{\"kind\":\"table\",\"id\":\"x\"},\"where\":{\"op\":\"blank\",\"col\":\"A\"}}", Json.Options);
        check("query spec wire shape", spec?.Where?.GetProperty("op").GetString() == "blank", "");
    }
}

public sealed class FakeProvider : ITableProvider
{
    readonly Table table; public FakeProvider(Table table) => this.table = table;
    public Table Get(TableRef r) => table;
    public IReadOnlyList<TableInfo> List() => Array.Empty<TableInfo>();
    public SourceInfo? SourceOf(int id) => null;
    public int SaveAsVersion(string reportTypeId, string period, Table data, string note) => throw new NotSupportedException();
}
