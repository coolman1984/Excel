// Lane B tests: filters, sorting, Distinct and timing. No Excel needed; an in-memory source stands in.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using ExcelApp;

/// <summary>A small in-memory <see cref="ICellSource"/> for the filter tests. Kind: 0 empty 1 number 2 text.</summary>
public sealed class MemoryCellSource : ICellSource
{
    readonly byte[][] _k;
    readonly double[][] _n;
    readonly string[][] _t;

    public int Rows { get; }
    public int Cols { get; }

    public MemoryCellSource(int rows, int cols)
    {
        Rows = rows;
        Cols = cols;
        _k = new byte[cols][];
        _n = new double[cols][];
        _t = new string[cols][];
        for (int c = 0; c < cols; c++)
        {
            _k[c] = new byte[rows];
            _n[c] = new double[rows];
            _t[c] = new string[rows];
        }
    }

    public void SetNumber(int row, int col, double v)
    {
        _k[col][row] = 1;
        _n[col][row] = v;
        _t[col][row] = v.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>Sets a number with the shown text given by the caller (used by the big timing source).</summary>
    public void SetNumber(int row, int col, double v, string shown)
    {
        _k[col][row] = 1;
        _n[col][row] = v;
        _t[col][row] = shown;
    }

    public void SetText(int row, int col, string s)
    {
        _k[col][row] = 2;
        _t[col][row] = s;
    }

    public void SetEmpty(int row, int col)
    {
        _k[col][row] = 0;
        _t[col][row] = "";
    }

    public void Read(int col, int fromRow, int count, string[] text, byte[] kind, double[] num)
    {
        for (int i = 0; i < count; i++)
        {
            int r = fromRow + i;
            kind[i] = _k[col][r];
            num[i] = _n[col][r];
            text[i] = _t[col][r] ?? "";
        }
    }
}

/// <summary>Tests for the filter and sort engine. See FilterEngine.cs.</summary>
public static class FilterTests
{
    /// <summary>Runs the checks. The thunk reports pass/fail: name, ok, detail.</summary>
    public static void Run(Action<string, bool, string> check)
    {
        ValuesFilter(check);
        NumberAndTextFilters(check);
        TwoColumnsAnd(check);
        SortAfterFilterKeepsHeaderFirst(check);
        Top10AndAverages(check);
        DistinctWithAnotherColumnsFilter(check);
        Timing(check);
    }

    static string Join(int[]? rows) => rows == null ? "null" : string.Join(",", rows);

    static void ValuesFilter(Action<string, bool, string> check)
    {
        var s = new MemoryCellSource(6, 2);
        s.SetText(0, 0, "Name");
        s.SetText(0, 1, "City");
        string[] names = { "Ann", "Bob", "Ann", "Cal", "Bob" };
        for (int i = 0; i < 5; i++) { s.SetText(1 + i, 0, names[i]); s.SetNumber(1 + i, 1, i + 1); }
        var fe = new FilterEngine(s) { HeaderRows = 1 };
        fe.SetCondition(0, new ValuesCondition(new HashSet<string> { "Ann" }, false));
        check("values: keeps exactly the chosen texts", Join(fe.Apply(CancellationToken.None)) == "0,1,3", Join(fe.Apply(CancellationToken.None)));

        var s2 = new MemoryCellSource(5, 1);
        s2.SetText(0, 0, "H");
        s2.SetText(1, 0, "a"); s2.SetEmpty(2, 0); s2.SetText(3, 0, "a"); s2.SetEmpty(4, 0);
        var fe2 = new FilterEngine(s2) { HeaderRows = 1 };
        fe2.SetCondition(0, new ValuesCondition(new HashSet<string> { "a" }, true));
        check("values: (Blanks) adds the empty cells", Join(fe2.Apply(CancellationToken.None)) == "0,1,2,3,4", Join(fe2.Apply(CancellationToken.None)));
        fe2.SetCondition(0, new ValuesCondition(new HashSet<string> { "a" }, false));
        check("values: without (Blanks) the empty cells are out", Join(fe2.Apply(CancellationToken.None)) == "0,1,3", Join(fe2.Apply(CancellationToken.None)));

        var fe3 = new FilterEngine(s) { HeaderRows = 1 };
        check("apply: no filter and no sort returns null (identity)", fe3.Apply(CancellationToken.None) == null, "");
    }

    static void NumberAndTextFilters(Action<string, bool, string> check)
    {
        var n = new MemoryCellSource(6, 1);
        n.SetText(0, 0, "N");
        for (int i = 1; i <= 5; i++) n.SetNumber(i, 0, i);
        var fn = new FilterEngine(n) { HeaderRows = 1 };
        fn.SetCondition(0, new NumberCondition("between", 2, 4));
        check("number: between 2 and 4", Join(fn.Apply(CancellationToken.None)) == "0,2,3,4", Join(fn.Apply(CancellationToken.None)));
        fn.SetCondition(0, new NumberCondition("gt", 3));
        check("number: greater than 3", Join(fn.Apply(CancellationToken.None)) == "0,4,5", Join(fn.Apply(CancellationToken.None)));
        fn.SetCondition(0, new NumberCondition("eq", 2));
        check("number: equals 2", Join(fn.Apply(CancellationToken.None)) == "0,2", Join(fn.Apply(CancellationToken.None)));

        var mix = new MemoryCellSource(3, 1);
        mix.SetText(0, 0, "H"); mix.SetNumber(1, 0, 5); mix.SetText(2, 0, "x");
        var fm = new FilterEngine(mix) { HeaderRows = 1 };
        fm.SetCondition(0, new NumberCondition("ne", 5));
        check("number: a text cell fails 'not equal' too", Join(fm.Apply(CancellationToken.None)) == "0", Join(fm.Apply(CancellationToken.None)));

        var t = new MemoryCellSource(6, 1);
        t.SetText(0, 0, "F");
        string[] words = { "apple", "banana", "cherry", "date", "elderberry" };
        for (int i = 0; i < 5; i++) t.SetText(1 + i, 0, words[i]);
        var ft = new FilterEngine(t) { HeaderRows = 1 };
        ft.SetCondition(0, new TextCondition("contains", "a"));
        check("text: contains 'a'", Join(ft.Apply(CancellationToken.None)) == "0,1,2,4", Join(ft.Apply(CancellationToken.None)));
        ft.SetCondition(0, new TextCondition("begins", "c"));
        check("text: begins with 'c'", Join(ft.Apply(CancellationToken.None)) == "0,3", Join(ft.Apply(CancellationToken.None)));
        ft.SetCondition(0, new TextCondition("ends", "y"));
        check("text: ends with 'y'", Join(ft.Apply(CancellationToken.None)) == "0,3,5", Join(ft.Apply(CancellationToken.None)));
    }

    static void TwoColumnsAnd(Action<string, bool, string> check)
    {
        var s = new MemoryCellSource(7, 2);
        s.SetText(0, 0, "N"); s.SetText(0, 1, "P");
        for (int i = 1; i <= 6; i++) { s.SetNumber(i, 0, i); s.SetText(i, 1, i % 2 == 0 ? "even" : "odd"); }
        var fe = new FilterEngine(s) { HeaderRows = 1 };
        fe.SetCondition(0, new NumberCondition("gt", 3));    // rows 4,5,6
        fe.SetCondition(1, new TextCondition("eq", "even")); // rows 2,4,6
        check("two columns: AND keeps only rows passing both", Join(fe.Apply(CancellationToken.None)) == "0,4,6", Join(fe.Apply(CancellationToken.None)));
    }

    static void SortAfterFilterKeepsHeaderFirst(Action<string, bool, string> check)
    {
        var s = new MemoryCellSource(5, 2);
        s.SetText(0, 0, "K"); s.SetText(0, 1, "V");
        string[] k = { "b", "a", "c", "a" };
        for (int i = 0; i < 4; i++) { s.SetText(1 + i, 0, k[i]); s.SetNumber(1 + i, 1, i + 1); }
        var fe = new FilterEngine(s) { HeaderRows = 1 };
        fe.SetCondition(0, new ValuesCondition(new HashSet<string> { "a", "c" }, false)); // rows 2,3,4
        fe.SetSort(new[] { (1, true) });
        var view = fe.Apply(CancellationToken.None);
        check("sort: header stays first after filter and sort", view != null && view[0] == 0, Join(view));
        check("sort: descending on the value column", Join(view) == "0,4,3,2", Join(view));

        var s2 = new MemoryCellSource(5, 1);
        s2.SetText(0, 0, "N");
        s2.SetNumber(1, 0, 3); s2.SetEmpty(2, 0); s2.SetNumber(3, 0, 1); s2.SetNumber(4, 0, 2);
        var fe2 = new FilterEngine(s2) { HeaderRows = 1 };
        fe2.SetSort(new[] { (0, false) });
        check("sort: numbers ascending, blanks last", Join(fe2.Apply(CancellationToken.None)) == "0,3,4,1,2", Join(fe2.Apply(CancellationToken.None)));
    }

    static bool Has(int[]? view, int row) => view != null && Array.IndexOf(view, row) >= 0;

    static void Top10AndAverages(Action<string, bool, string> check)
    {
        var s = new MemoryCellSource(21, 1);
        s.SetText(0, 0, "N");
        s.SetNumber(1, 0, 100);                          // one big value
        for (int i = 2; i <= 20; i++) s.SetNumber(i, 0, i - 1); // 1..19
        // Average is (100 + 190) / 20 = 14.5.
        var fe = new FilterEngine(s) { HeaderRows = 1 };

        fe.SetCondition(0, new NumberCondition("aboveAverage", 0));
        var above = fe.Apply(CancellationToken.None);
        check("average: aboveAverage keeps values over 14.5", above != null && above.Length == 7, above == null ? "null" : (above.Length - 1).ToString());

        fe.SetCondition(0, new NumberCondition("belowAverage", 0));
        var below = fe.Apply(CancellationToken.None);
        check("average: belowAverage keeps values under 14.5", below != null && below.Length == 15, below == null ? "null" : (below.Length - 1).ToString());

        fe.SetCondition(0, new NumberCondition("top10", 10));
        var top = fe.Apply(CancellationToken.None);
        check("top10: keeps the ten biggest (10 rows)", top != null && top.Length == 11, top == null ? "null" : (top.Length - 1).ToString());
        check("top10: contains the 100 and the 11, not the 10", Has(top, 1) && Has(top, 12) && !Has(top, 11), Join(top));
    }

    static void DistinctWithAnotherColumnsFilter(Action<string, bool, string> check)
    {
        var s = new MemoryCellSource(6, 2);
        s.SetText(0, 0, "G"); s.SetText(0, 1, "V");
        s.SetText(1, 0, "A"); s.SetText(1, 1, "x");
        s.SetText(2, 0, "A"); s.SetText(2, 1, "y");
        s.SetText(3, 0, "B"); s.SetText(3, 1, "x");
        s.SetText(4, 0, "A"); s.SetText(4, 1, "x");
        s.SetText(5, 0, "B"); s.SetText(5, 1, "z");
        var fe = new FilterEngine(s) { HeaderRows = 1 };

        var all = fe.Distinct(0, null);
        check("distinct: counts over all rows", all.Values.Count == 2 && all.Values[0] == ("A", 3) && all.Values[1] == ("B", 2), string.Join(";", all.Values));

        fe.SetCondition(0, new TextCondition("eq", "A"));
        var other = fe.Distinct(1, null); // among rows 1,2,4: x,y,x
        check("distinct: obeys another column's filter", other.Values.Count == 2 && other.Values[0] == ("x", 2) && other.Values[1] == ("y", 1), string.Join(";", other.Values));

        var self = fe.Distinct(0, null);  // the column's own condition is ignored
        check("distinct: ignores the column's own condition", self.Values.Count == 2 && self.Values[0] == ("A", 3), string.Join(";", self.Values));

        var n = new MemoryCellSource(6, 1);
        n.SetText(0, 0, "N");
        n.SetNumber(1, 0, 5); n.SetNumber(2, 0, 1); n.SetEmpty(3, 0); n.SetNumber(4, 0, 5); n.SetNumber(5, 0, 3);
        var fan = new FilterEngine(n) { HeaderRows = 1 };
        var v = fan.Distinct(0, null);
        check("distinct: numbers ascending, blanks counted apart",
            v.Values.Count == 3 && v.Values[0].Text == "1" && v.Values[1].Text == "3" && v.Values[2].Text == "5" && v.Blanks == 1,
            string.Join(";", v.Values) + " blanks=" + v.Blanks);
        var v2 = fan.Distinct(0, "5");
        check("distinct: search keeps only matching texts", v2.Values.Count == 1 && v2.Values[0] == ("5", 2), string.Join(";", v2.Values));
    }

    static void Timing(Action<string, bool, string> check)
    {
        const int N = 1_000_000;
        var s = new MemoryCellSource(N, 4);
        var thousand = new string[1000];
        for (int i = 0; i < 1000; i++) thousand[i] = i.ToString(CultureInfo.InvariantCulture);
        var cat = new string[50];
        for (int i = 0; i < 50; i++) cat[i] = "cat" + i;
        var vv = new string[5000];
        for (int i = 0; i < 5000; i++) vv[i] = "v" + i;
        for (int r = 0; r < N; r++)
        {
            s.SetNumber(r, 0, r % 1000, thousand[r % 1000]);
            s.SetText(r, 1, cat[r % 50]);
            s.SetNumber(r, 2, r, "");
            s.SetText(r, 3, vv[r % 5000]);
        }

        var fe = new FilterEngine(s) { HeaderRows = 1 };
        fe.SetCondition(0, new NumberCondition("gt", 500));
        fe.Apply(CancellationToken.None); // warm up the JIT and the code paths first
        var sw = Stopwatch.StartNew();
        var view = fe.Apply(CancellationToken.None);
        sw.Stop();
        check($"timing: Apply on 1,000,000 rows under 1 s (took {sw.ElapsedMilliseconds} ms)", view != null && sw.ElapsedMilliseconds < 1000, sw.ElapsedMilliseconds + " ms");

        var fe2 = new FilterEngine(s) { HeaderRows = 1 };
        fe2.Distinct(3, "v1"); // warm up
        var swc = Stopwatch.StartNew();
        var d = fe2.Distinct(3, "v1");
        swc.Stop();
        check($"timing: Distinct on 1,000,000 rows under 1 s (took {swc.ElapsedMilliseconds} ms)", d.Values.Count > 0 && swc.ElapsedMilliseconds < 1000, swc.ElapsedMilliseconds + " ms");
    }
}


