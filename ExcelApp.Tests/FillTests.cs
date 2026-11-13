// The fill handle's series rules (ExcelApp.FillSeries) against Excel's AutoFill behaviour, without a window or Excel.
using ExcelApp;

public static class FillTests
{
    static CellEdit T(string s) => new(EditKind.Text, Text: s);
    static CellEdit N(double d) => new(EditKind.Number, Num: d);

    public static void Run(Action<string, bool, string> check)
    {
        string Text(FillSeries s, int t) => s.At(t)?.Text ?? "(copy)";
        double Num(FillSeries s, int t) => s.At(t)?.Num ?? double.NaN;
        FillSeries Of(params CellEdit[] items) => FillSeries.Of(items, items.Select(_ => "General").ToArray());

        var one = Of(T("Item 1"));
        check("fill series: one text cell counts up", Text(one, 1) == "Item 2" && Text(one, 3) == "Item 4", Text(one, 1));
        var two = Of(T("item-1"), T("item-2"));
        check("fill series: two numbered text cells continue the step", Text(two, 2) == "item-3" && Text(two, 3) == "item-4", Text(two, 2) + "," + Text(two, 3));
        var odd = Of(T("Week 1"), T("Week 3"));
        check("fill series: a text step of two", Text(odd, 2) == "Week 5" && Text(odd, 3) == "Week 7", Text(odd, 2) + "," + Text(odd, 3));
        var zeros = Of(T("A001"), T("A002"));
        check("fill series: leading zeros kept", Text(zeros, 2) == "A003" && Text(zeros, 9) == "A010", Text(zeros, 2));
        var down = Of(T("Item 3"), T("Item 2"));
        check("fill series: counting down goes past zero as Excel does", Text(down, 2) == "Item 1" && Text(down, 3) == "Item 0" && Text(down, 4) == "Item 1", Text(down, 3) + "," + Text(down, 4));
        var mixed = Of(T("Text1"), T("textA"));
        check("fill series: mixed cells, the numbered one counts per cycle", Text(mixed, 2) == "Text2" && Text(mixed, 3) == "(copy)" && Text(mixed, 4) == "Text3", Text(mixed, 2) + "," + Text(mixed, 3) + "," + Text(mixed, 4));
        var prefixes = Of(T("a1"), T("b1"));
        check("fill series: different prefixes each count per cycle", Text(prefixes, 2) == "a2" && Text(prefixes, 3) == "b2", Text(prefixes, 2) + "," + Text(prefixes, 3));
        var huge = Of(T("id" + new string('9', 15)));
        check("fill series: a long number still counts", Text(huge, 1) == "id1000000000000000", Text(huge, 1));
        var plain = Of(T("north"), T("south"));
        check("fill series: plain text repeats", Text(plain, 2) == "(copy)", Text(plain, 2));
        var days = Of(T("Mon"), T("Wed"));
        check("fill series: day names step", Text(days, 2) == "Fri" && Text(days, 3) == "Sun", Text(days, 2));
        var months = Of(T("JANUARY"));
        check("fill series: month names keep upper case", Text(months, 1) == "FEBRUARY" && Text(months, 12) == "JANUARY", Text(months, 1));
        var nums = Of(N(1), N(3));
        check("fill series: two numbers make a trend", Num(nums, 2) == 5 && Num(nums, 3) == 7, Num(nums, 2).ToString());
        var single = Of(N(4));
        check("fill series: one plain number is copied", single.At(1) == null, "");
        var date = FillSeries.Of(new[] { N(45000) }, new[] { "dd/mm/yyyy" });
        check("fill series: one date counts by a day", Num(date, 2) == 45002, "");
        var fit = Of(N(1), N(2), N(4));
        check("fill series: three numbers make a best-fit line", Math.Abs(Num(fit, 3) - 5.33333333333333) < 1e-9, Num(fit, 3).ToString());
    }
}
