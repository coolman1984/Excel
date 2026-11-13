// The editor's formula engine without Excel: functions against Excel's own answers, reference moving, and bracket rules.
using System.Text.RegularExpressions;
using ExcelApp.Formula;

public static class CalcTests
{
    sealed class Grid : IFormulaGrid
    {
        public readonly Dictionary<(int, int, int), FValue> Cells = new();
        public FValue Get(int sheet, int row, int col) => Cells.TryGetValue((sheet, row, col), out var v) ? v : FValue.Blank;
        public int SheetIndex(string name) => name.Equals("Data", StringComparison.OrdinalIgnoreCase) ? 1 : name.Equals("Sheet1", StringComparison.OrdinalIgnoreCase) ? 0 : -1;
        public (int Rows, int Cols) Extent(int sheet) => (20, 10);
        public void Set(string a1, object v, int sheet = 0)
        {
            var m = Regex.Match(a1, @"^([A-Z]+)(\d+)$"); int col = 0; foreach (char c in m.Groups[1].Value) col = col * 26 + c - 'A' + 1;
            Cells[(sheet, int.Parse(m.Groups[2].Value) - 1, col - 1)] = v switch { double d => FValue.Num(d), int i => FValue.Num(i), bool b => FValue.Bool(b), string s when s.StartsWith('#') => FValue.Error(s), string s => FValue.Str(s), _ => FValue.Blank };
        }
    }

    public static void Run(Action<string, bool, string> check)
    {
        var g = new Grid();
        // A1:A5 = 1..5, B1:B5 = names, C1:C5 = 10,20,30,40,50
        string[] names = { "apple", "banana", "cherry", "date", "elder" };
        for (int i = 1; i <= 5; i++) { g.Set("A" + i, i); g.Set("B" + i, names[i - 1]); g.Set("C" + i, i * 10); }
        g.Set("A1", 7, 1);   // Data!A1
        var ev = new Evaluator(g) { Row = 9, Col = 3 };   // the formula sits in D10

        void Num(string f, double want, double tol = 1e-9)
        {
            FValue v; try { v = ev.Evaluate(f); } catch (Exception e) { check("calc " + f, false, e.GetType().Name); return; }
            check("calc " + f, v.Kind == ValueKind.Number && Math.Abs(v.Number - want) <= tol, $"{v.Kind} {v.Number} {v.Text}");
        }
        void Str(string f, string want) { var v = ev.Evaluate(f); check("calc " + f, v.Kind == ValueKind.Text && v.Text == want, $"{v.Kind} {v.Text}"); }
        void Bool(string f, bool want) { var v = ev.Evaluate(f); check("calc " + f, v.Kind == ValueKind.Bool && (v.Number != 0) == want, $"{v.Kind} {v.Number}"); }
        void Err(string f, string want) { var v = ev.Evaluate(f); check("calc " + f, v.Kind == ValueKind.Error && v.Text == want, $"{v.Kind} {v.Text}"); }

        // arithmetic and precedence, as Excel
        Num("=1+2*3", 7); Num("=-2^2", 4); Num("=2^3^2", 64); Num("=50%", 0.5); Num("=(1+2)*3", 9); Num("=\"3\"+1", 4);
        Str("=\"a\"&1&TRUE", "a1TRUE"); Err("=1/0", "#DIV/0!"); Err("=\"x\"+1", "#VALUE!");
        // maths
        Num("=SUM(A1:A5)", 15); Num("=SUM(A:A)", 15); Num("=PRODUCT(A1:A4)", 24); Num("=SUMPRODUCT(A1:A5,C1:C5)", 550); Num("=SUMSQ(A1:A3)", 14);
        Num("=ROUND(2.345,2)", 2.35); Num("=ROUND(-2.5,0)", -3); Num("=ROUNDUP(2.341,2)", 2.35); Num("=ROUNDDOWN(-2.349,2)", -2.34); Num("=TRUNC(-4.7)", -4);
        Num("=MOD(-3,2)", 1); Num("=INT(-1.5)", -2); Num("=SQRT(16)", 4); Num("=POWER(2,10)", 1024); Num("=LOG(8,2)", 3); Num("=LOG10(1000)", 3); Num("=LN(EXP(2))", 2);
        Num("=CEILING(2.5,1)", 3); Num("=CEILING(-2.5,-2)", -4); Num("=CEILING(-2.5,2)", -2); Num("=FLOOR(-2.5,2)", -4); Num("=FLOOR(2.5,1)", 2); Num("=FLOOR.MATH(-2.5)", -3); Num("=CEILING.MATH(-2.5)", -2);
        Num("=MROUND(10,3)", 9); Num("=EVEN(1.5)", 2); Num("=ODD(2)", 3); Num("=ODD(-2)", -3); Num("=FACT(5)", 120); Num("=GCD(12,18)", 6); Num("=LCM(4,6)", 12);
        Num("=QUOTIENT(7,2)", 3); Num("=SIGN(-4)", -1); Num("=ABS(-3)", 3); Num("=DEGREES(PI())", 180); Num("=SUBTOTAL(9,A1:A5)", 15); Num("=SUBTOTAL(101,A1:A5)", 3);
        // statistics
        Num("=AVERAGE(A1:A5)", 3); Num("=MEDIAN(1,3,2,4)", 2.5); Num("=MODE(1,2,2,3)", 2); Num("=STDEV.S(A1:A5)", 1.5811388300841898); Num("=STDEV.P(A1:A5)", 1.4142135623730951);
        Num("=VAR.S(A1:A5)", 2.5); Num("=LARGE(C1:C5,2)", 40); Num("=SMALL(C1:C5,2)", 20); Num("=RANK(3,A1:A5)", 3); Num("=RANK(3,A1:A5,1)", 3); Num("=PERCENTILE(A1:A5,0.25)", 2); Num("=QUARTILE(A1:A5,3)", 4);
        Num("=COUNT(A1:C5)", 10); Num("=COUNTA(A1:C5)", 15); Num("=COUNTBLANK(D1:D3)", 3); Num("=COUNTIF(A1:A5,\">2\")", 3); Num("=COUNTIFS(A1:A5,\">1\",C1:C5,\"<50\")", 3);
        Num("=SUMIF(B1:B5,\"*e*\",C1:C5)", 10 + 30 + 40 + 50); Num("=AVERAGEIF(A1:A5,\">=4\")", 4.5); Num("=MAXIFS(C1:C5,A1:A5,\"<4\")", 30); Num("=MINIFS(C1:C5,A1:A5,\">2\")", 30);
        Num("=MAX(A1:A5)", 5); Num("=MIN(C1:C5)", 10); Num("=COUNTUNIQUE(1,1,2,\"a\",\"A\")", 3);
        // logic and information
        Num("=IF(A1>0,10,20)", 10); Num("=IFERROR(1/0,99)", 99); Num("=IFS(A1>5,1,A1>0,2)", 2); Num("=CHOOSE(2,10,20,30)", 20); Num("=SWITCH(2,1,\"a\",2,5,9)", 5);
        Bool("=AND(TRUE,1)", true); Bool("=OR(FALSE,0)", false); Bool("=XOR(TRUE,TRUE,TRUE)", true); Bool("=NOT(0)", true); Bool("=ISBLANK(D1)", true); Bool("=ISNUMBER(A1)", true);
        Bool("=ISTEXT(B1)", true); Bool("=ISERROR(1/0)", true); Bool("=ISNA(NA())", true); Bool("=ISERR(NA())", false); Bool("=ISEVEN(4)", true); Bool("=ISODD(4)", false); Bool("=TRUE()", true);
        Num("=ROW()", 10); Num("=COLUMN()", 4); Num("=ROW(C7)", 7); Num("=ROWS(A1:B5)", 5); Num("=COLUMNS(A1:C1)", 3); Num("=Data!A1*2", 14); Num("='Data'!A1+1", 8);
        // lookup
        Num("=VLOOKUP(\"cherry\",B1:C5,2,FALSE)", 30); Num("=VLOOKUP(3.5,A1:C5,3)", 30); Num("=HLOOKUP(10,A1:C2,2,FALSE)", 20); Num("=INDEX(C1:C5,4)", 40); Num("=MATCH(\"date\",B1:B5,0)", 4);
        Num("=XLOOKUP(\"elder\",B1:B5,C1:C5)", 50); Str("=XLOOKUP(\"zzz\",B1:B5,B1:B5,\"none\")", "none"); Err("=VLOOKUP(\"zzz\",B1:C5,2,FALSE)", "#N/A");
        // text
        Str("=LEFT(\"Excel\",2)", "Ex"); Str("=RIGHT(\"Excel\",3)", "cel"); Str("=MID(\"Excel\",2,3)", "xce"); Num("=LEN(\"Excel\")", 5); Str("=TRIM(\"  a   b \")", "a b");
        Str("=UPPER(\"ab\")", "AB"); Str("=PROPER(\"hello wORLD\")", "Hello World"); Num("=FIND(\"l\",\"hello\")", 3); Num("=SEARCH(\"L*O\",\"hello\")", 3); Err("=FIND(\"L\",\"hello\")", "#VALUE!");
        Str("=SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "a+b+c"); Str("=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "a-b+c"); Str("=REPLACE(\"abcdef\",2,3,\"X\")", "aXef"); Str("=REPT(\"ab\",3)", "ababab");
        Bool("=EXACT(\"a\",\"A\")", false); Str("=CHAR(65)", "A"); Num("=CODE(\"A\")", 65); Str("=CONCAT(\"a\",B1)", "aapple"); Str("=TEXTJOIN(\",\",TRUE,\"a\",\"\",\"b\")", "a,b");
        Str("=TEXTBEFORE(\"a.b.c\",\".\")", "a"); Str("=TEXTAFTER(\"a.b.c\",\".\",-1)", "c"); Str("=FIXED(1234.567,1)", "1,234.6"); Num("=VALUE(\"12.5\")", 12.5); Num("=NUMBERVALUE(\"1,234.5\",\".\",\",\")", 1234.5);
        Str("=TEXT(0.5,\"0%\")", "50%");
        // dates (serials as Excel: 45292 is 2024-01-01)
        Num("=DATE(2024,1,1)", 45292); Num("=DATE(2024,13,1)", 45658); Num("=YEAR(45292)", 2024); Num("=MONTH(45292)", 1); Num("=DAY(45292)", 1); Num("=EOMONTH(45292,1)", 45351);
        Num("=EDATE(45322,1)", 45351); Num("=WEEKDAY(45292)", 2); Num("=DATEDIF(DATE(2020,5,15),DATE(2024,5,14),\"Y\")", 3); Num("=DATEDIF(DATE(2024,1,31),DATE(2024,3,1),\"M\")", 1);
        Num("=DAYS(DATE(2024,3,1),DATE(2024,2,1))", 29); Num("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,12))", 10); Num("=WORKDAY(DATE(2024,1,5),1)", 45299);
        Num("=TIME(12,30,0)", 0.5208333333333334); Num("=HOUR(0.75)", 18); Num("=MINUTE(TIME(1,45,0))", 45); Num("=SECOND(TIME(1,2,3))", 3); Num("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))", 0.5);
        // finance (Excel's answers)
        Num("=PMT(0.05/12,60,10000)", -188.71233644010877, 1e-8); Num("=FV(0.06/12,10,-200,-500,1)", 2581.4033740601, 1e-6); Num("=PV(0.08/12,240,500)", -59777.14585118, 1e-6);
        Num("=NPV(0.1,-10000,3000,4200,6800)", 1188.4434123352, 1e-6); Num("=NPER(0.01,-100,1000)", 10.588644459, 1e-6); Num("=IPMT(0.1/12,1,36,8000)", -66.6666666667, 1e-6); Num("=PPMT(0.1/12,1,24,2000)", -75.6231860084, 1e-6);
        // unknown names are not guessed
        check("calc: an unknown function is reported, not guessed", ev.Evaluate("=NOSUCH(1)").Kind == ValueKind.Unsupported, "");

        // ---- references move the way Excel moves them
        void Shift(string f, int dr, int dc, string want) { string got = FormulaShift.Shift(f, dr, dc); check($"shift {f} by {dr},{dc}", got == want, got); }
        Shift("=A1*2", 1, 0, "=A2*2"); Shift("=$A$1+A$1+$A1", 2, 3, "=$A$1+D$1+$A3"); Shift("=SUM(B2:B4)", 0, 1, "=SUM(C2:C4)"); Shift("=SUM(A:A)", 0, 2, "=SUM(C:C)");
        Shift("=Data!A1+'My Sheet'!B2", 1, 1, "=Data!B2+'My Sheet'!C3"); Shift("=\"A1\"&A1", 1, 0, "=\"A1\"&A2"); Shift("=A1", -1, 0, "=#REF!"); Shift("=LOG10(A1)", 1, 0, "=LOG10(A2)");
        Shift("=ROUND(A1,2)", 0, 0, "=ROUND(A1,2)");

        // ---- Excel's bracket rules on entry
        check("complete: a missing closing bracket is added", Calc.Complete("=SUM(A1:A3") == "=SUM(A1:A3)" && Calc.Complete("=IF(A1>0,ROUND(A1,2") == "=IF(A1>0,ROUND(A1,2))", "");
        check("problem: an extra closing bracket is refused", Calc.Problem("=SUM(1))") != null && Calc.Problem("=SUM(1)") == null, "");
        check("problem: an open text value is refused", Calc.Problem("=\"abc") != null && Calc.Problem("=\"a(b\"") == null, "");

        // ---- the interface's function list names only functions the engine calculates
        var root = AppContext.BaseDirectory; string? ts = null;
        for (var d = new DirectoryInfo(root); d != null; d = d.Parent) { var p = Path.Combine(d.FullName, "ExcelApp", "ui", "src", "functions.ts"); if (File.Exists(p)) { ts = p; break; } }
        if (ts == null) check("functions.ts found", false, root);
        else
        {
            var listed = Regex.Matches(File.ReadAllText(ts), @"F\(""[^""]+"", ""([A-Z0-9.]+)""").Select(m => m.Groups[1].Value).ToHashSet();
            var engine = Evaluator.Functions.ToHashSet();
            var missing = listed.Where(n => !engine.Contains(n)).ToList(); var unlisted = engine.Where(n => !listed.Contains(n)).ToList();
            check("functions: every listed function is calculated", missing.Count == 0, string.Join(",", missing));
            check("functions: every calculated function is listed", unlisted.Count == 0, string.Join(",", unlisted));
        }
    }
}
