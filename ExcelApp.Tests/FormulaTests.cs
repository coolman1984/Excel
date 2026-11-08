using ExcelApp.Formula;
using System.Runtime.InteropServices;

public static class FormulaTests
{
    public static void Run(Action<string, bool, string> check, bool offline)
    {
        var grid = new Grid(); var evaluator = new Evaluator(grid); grid.Cells[new Cell(0, 0, 0)] = FValue.Num(5); grid.Cells[new Cell(0, 1, 0)] = FValue.Num(7); grid.Cells[new Cell(0, 0, 1)] = FValue.Str("Alpha"); grid.Cells[new Cell(0, 1, 1)] = FValue.Str("Beta");
        void Num(string formula, double expected) { var value = evaluator.Evaluate(formula); check("formula " + formula, value.Kind == ValueKind.Number && Math.Abs(value.Number - expected) < 1e-9, ""); }
        Num("=A1+A2", 12); Num("=SUM(A1:A2)", 12); Num("=A2/A1", 1.4); Num("=IF(A1>0,8,9)", 8); Num("=ROUND(-1.25,1)", -1.3); Num("=MOD(-3,2)", 1); Num("=50%", .5);
        Num("=COUNTIF(A1:A2,\">5\")", 1); Num("=SUMIF(A1:A2,\">5\")", 7); Num("=AVERAGEIF(A1:A2,\">5\")", 7); Num("=INDEX(A1:A2,2)", 7); Num("=MATCH(7,A1:A2,0)", 2); Num("=XLOOKUP(5,A1:A2,A1:A2)", 5);
        Num("=YEAR(DATE(2026,11,8))", 2026); Num("=MONTH(DATE(2026,11,8))", 11); Num("=DAY(EOMONTH(DATE(2026,2,1),0))", 28);
        check("formula wildcard", evaluator.Evaluate("=COUNTIF(B1:B2,\"a*\")").Number == 1, "");
        check("formula left and mid", evaluator.Evaluate("=LEFT(\"hello\",2)").Text == "he" && evaluator.Evaluate("=MID(\"hello\",2,3)").Text == "ell", "");
        Num("=IFS(A1<0,1,A1>0,2)", 2); Num("=SUMIFS(A1:A2,A1:A2,\">5\")", 7); Num("=COUNTIFS(A1:A2,\">4\",A1:A2,\"<7\")", 1);
        Num("=VLOOKUP(7,A1:B2,1,FALSE)", 7); Num("=XLOOKUP(6,A1:A2,A1:A2,0,-1)", 5); Num("=Sheet1!$A$1", 5);
        check("unsupported ignores error argument", evaluator.Evaluate("=UNKNOWN(1/0)").Kind == ValueKind.Unsupported, "");
        check("formula unsupported", evaluator.Evaluate("=UNRECOGNIZED(1)").Kind == ValueKind.Unsupported, "");
        check("formula divide zero", evaluator.Evaluate("=1/0").Text == "#DIV/0!", "");
        check("formula text", evaluator.Evaluate("=UPPER(\"abc\")").Text == "ABC", "");
        for (int n = 1; n <= 15; n++)
        {
            Num($"={n}+3", n + 3); Num($"={n}-3", n - 3); Num($"={n}*3", n * 3); Num($"={n}/2", n / 2.0); Num($"={n}^2", n * n);
            Num($"=ABS(-{n})", n); Num($"=SUM({n},2)", n + 2); Num($"=MAX({n},2)", Math.Max(n, 2)); Num($"=MIN({n},2)", Math.Min(n, 2)); Num($"=AVERAGE({n},2)", (n + 2) / 2.0);
        }
        var graph = new DependencyGraph(); var input = new Cell(0, 0, 0); var first = new Cell(0, 0, 1); var second = new Cell(0, 0, 2);
        graph.Set(first, new[] { new CellRange(0, 0, 0, 0, 0) }); graph.Set(second, new[] { new CellRange(0, 0, 0, 1, 1) }); var order = graph.Recalculate(new[] { input });
        check("formula dependency order", order.Ordered.SequenceEqual(new[] { first, second }) && order.Circular.Count == 0, "");
        var cyclic = new DependencyGraph(); var a = new Cell(0, 1, 0); var b = new Cell(0, 2, 0); var downstream = new Cell(0, 3, 0);
        cyclic.Set(a, new[] { new CellRange(0, 2, 2, 0, 0) }); cyclic.Set(b, new[] { new CellRange(0, 1, 1, 0, 0) }); cyclic.Set(downstream, new[] { new CellRange(0, 2, 2, 0, 0) }); var cycle = cyclic.Recalculate(new[] { a });
        check("formula cycle excludes downstream", cycle.Circular.SetEquals(new[] { a, b }) && cycle.Ordered.Contains(downstream), "");
        var large = new DependencyGraph(); for (int i = 0; i < 100_000; i++) large.Set(new Cell(0, i, 1), new[] { new CellRange(0, 0, 0, 0, 0) });
        var compiled = new Parser("=A1+1").Parse(); var watch = System.Diagnostics.Stopwatch.StartNew(); var recalc = large.Recalculate(new[] { input }); double total = 0; foreach (var _ in recalc.Ordered) total += evaluator.Evaluate(compiled).Number; watch.Stop(); Console.WriteLine($"SPEED formula recalculation 100k: {watch.Elapsed.TotalSeconds:F3}s (target <1s)");
        check("formula recalculation speed 100k", recalc.Ordered.Count == 100_000 && total == 600_000 && watch.Elapsed.TotalSeconds < 1, "");
        if (!offline) Live(check);
    }
    sealed class Grid : IFormulaGrid
    {
        public Dictionary<Cell, FValue> Cells { get; } = new();
        public FValue Get(int sheet, int row, int col) => Cells.TryGetValue(new Cell(sheet, row, col), out var value) ? value : FValue.Blank;
        public int SheetIndex(string name) => name.Equals("Sheet1", StringComparison.OrdinalIgnoreCase) ? 0 : -1;
    }
    static void Live(Action<string, bool, string> check)
    {
        dynamic? app = null, books = null, book = null, sheets = null, sheet = null, cell = null;
        try
        {
            var type = Type.GetTypeFromProgID("Excel.Application") ?? throw new InvalidOperationException(); app = Activator.CreateInstance(type)!; app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false; app.AutomationSecurity = 3;
            books = app.Workbooks; book = books.Add(); sheets = book.Worksheets; sheet = sheets.Item[1]; cell = sheet.Cells[1, 1]; cell.Value2 = 5; Release(cell); cell = null;
            var grid = new Grid(); grid.Cells[new Cell(0, 0, 0)] = FValue.Num(5); var evaluator = new Evaluator(grid); int mismatches = 0;
            for (int i = 1; i <= 60; i++) { int n = (i - 1) / 12 + 1; string formula = ((i - 1) % 12) switch { 0 => $"=A1+{n}", 1 => $"=A1-{n}", 2 => $"=A1*{n}", 3 => $"=A1/{n}", 4 => $"=SUM(A1,{n})", 5 => $"=AVERAGE(A1,{n})", 6 => $"=MAX(A1,{n})", 7 => $"=MIN(A1,{n})", 8 => $"=ROUND(A1/{n},2)", 9 => $"=IF(A1>{n},A1,{n})", 10 => $"=MOD(A1,{n})", _ => $"=ABS(-{n})" }; cell = sheet.Cells[i, 2]; cell.Formula = formula; double excel = Convert.ToDouble(cell.Value2); var own = evaluator.Evaluate(formula); if (own.Kind != ValueKind.Number || Math.Abs(excel - own.Number) > 1e-9) { mismatches++; check("live formula mismatch " + i, false, ""); } Release(cell); cell = null; }
            check("live Excel 60 formula comparison", mismatches == 0, "");
            check("live formula reader finds 60", FormulaReader.Read(book).Count == 60, "");
        }
        finally { Release(cell); Release(sheet); Release(sheets); if (book != null) { try { book.Close(false); } catch (COMException) { } } Release(book); Release(books); if (app != null) { try { app.Quit(); } catch (COMException) { } } Release(app); }
    }
    static void Release(object? value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
}
