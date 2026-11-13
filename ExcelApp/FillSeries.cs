// The fill handle's series rules, kept apart from Session so they can be tested without a window or Excel.
using System.Globalization;
namespace ExcelApp;

/// <summary>How Excel's fill handle continues a run of cells: numbers as a straight-line trend (two or more),
/// a single date by one day, text ending in a number by one, and day and month names in order.</summary>
public sealed class FillSeries
{
    Func<int, CellEdit?>? at;
    public CellEdit? At(int t) => at?.Invoke(t);
    static readonly string[][] Names =
    {
        CultureInfo.InvariantCulture.DateTimeFormat.DayNames, CultureInfo.InvariantCulture.DateTimeFormat.AbbreviatedDayNames,
        CultureInfo.InvariantCulture.DateTimeFormat.MonthNames.Take(12).ToArray(), CultureInfo.InvariantCulture.DateTimeFormat.AbbreviatedMonthNames.Take(12).ToArray(),
    };
    static int Mod(int a, int n) => ((a % n) + n) % n;
    static bool IsDate(string f) { string x = System.Text.RegularExpressions.Regex.Replace(f, "\"[^\"]*\"|\\[[^\\]]*\\]", ""); return x.IndexOfAny(new[] { 'd', 'D', 'y', 'Y' }) >= 0; }

    public static FillSeries Of(CellEdit?[] items, string[] formats)
    {
        var s = new FillSeries(); int n = items.Length;
        if (n == 0 || items.Any(x => x == null || x.Kind is EditKind.Formula or EditKind.Clear)) return s;
        if (items.All(x => x!.Kind == EditKind.Number))
        {
            if (n == 1)
            {
                if (!IsDate(formats[0])) return s;   // one plain number is copied, as in Excel
                double v0 = items[0]!.Num; s.at = t => new CellEdit(EditKind.Number, Num: v0 + t);
                return s;
            }
            double mx = (n - 1) / 2.0, my = items.Average(x => x!.Num), sxy = 0, sxx = 0;
            for (int i = 0; i < n; i++) { sxy += (i - mx) * (items[i]!.Num - my); sxx += (i - mx) * (i - mx); }
            double slope = sxx == 0 ? 0 : sxy / sxx, icept = my - slope * mx;
            s.at = t => new CellEdit(EditKind.Number, Num: Cells.Excel15(icept + slope * t));
            return s;
        }
        if (items.All(x => x!.Kind == EditKind.Text))
        {
            var texts = items.Select(x => x!.Text ?? "").ToArray();
            foreach (var list in Names)
            {
                int[] idx = texts.Select(t => Array.FindIndex(list, nm => string.Equals(nm, t, StringComparison.OrdinalIgnoreCase))).ToArray();
                if (idx.Any(i => i < 0)) continue;
                int step = n >= 2 ? idx[1] - idx[0] : 1;
                bool upper = texts[0].Length > 1 && texts[0] == texts[0].ToUpperInvariant(), lower = texts[0] == texts[0].ToLowerInvariant();
                s.at = t =>
                {
                    string nm = list[Mod(idx[0] + step * t, list.Length)];
                    return new CellEdit(EditKind.Text, Text: upper ? nm.ToUpperInvariant() : lower ? nm.ToLowerInvariant() : nm);
                };
                return s;
            }
            var parts = texts.Select(t => System.Text.RegularExpressions.Regex.Match(t, @"^(.*?)(\d{1,15})$")).ToArray();
            if (!parts.Any(m => m.Success)) return s;
            string Shown(System.Text.RegularExpressions.Match m, long num) =>
                m.Groups[1].Value + Math.Abs(num).ToString(CultureInfo.InvariantCulture).PadLeft(m.Groups[2].Value.Length, '0');
            long[] nums = parts.Select(m => m.Success ? long.Parse(m.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture) : 0).ToArray();
            // Two or more numbered cells with one prefix and an even step continue that step ("Week 1, Week 3" gives "Week 5");
            // below zero Excel counts back up ("Item 1, Item 0, Item 1").
            if (n >= 2 && parts.All(m => m.Success) && parts.All(m => string.Equals(m.Groups[1].Value, parts[0].Groups[1].Value, StringComparison.OrdinalIgnoreCase)))
            {
                long step = nums[1] - nums[0];
                if (Enumerable.Range(1, n - 1).All(i => nums[i] - nums[i - 1] == step))
                {
                    s.at = t => new CellEdit(EditKind.Text, Text: Shown(parts[0], nums[0] + step * t));
                    return s;
                }
            }
            // Otherwise each numbered cell counts up by one per repeat of the source and the other cells repeat ("Text1, textA" gives "Text2, textA").
            s.at = t =>
            {
                int k = Mod(t, n), cycle = (t - k) / n;
                return parts[k].Success ? new CellEdit(EditKind.Text, Text: Shown(parts[k], nums[k] + cycle)) : null;
            };
        }
        return s;
    }
}
