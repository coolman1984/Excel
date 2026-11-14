// What a chart plots from a range, decided the way Excel does: a text header row names the series, a first column of
// text or dates holds the categories (also when the top-left cell is empty), every other column is a series.
using System.Globalization;

namespace ExcelApp;

public sealed record ChartSeries(string Name, double?[] Values);
/// <summary>X: the scatter chart's numeric x values (null for other charts, or when the first column is text).</summary>
public sealed record ChartPoints(string[] Categories, ChartSeries[] Series, double?[]? X, bool Truncated);

public static class ChartData
{
    public const int MaxPoints = 1000, MaxSeries = 8;
    public static readonly string[] Types = { "column", "bar", "line", "area", "pie", "scatter" };

    /// <summary>Kind: 0 empty, 1 number, 2 text, 3 bool, 4 error (ExcelViewer.Core.CellKind); Text: as shown.</summary>
    public static ChartPoints Read(int rows, int cols, Func<int, int, (byte Kind, double Num, string Text)> cell, string type)
    {
        if (rows <= 0 || cols <= 0) return new(Array.Empty<string>(), Array.Empty<ChartSeries>(), null, false);
        bool Number(int r, int c) { var v = cell(r, c); return v.Kind == 1 && double.TryParse(v.Text.Replace(",", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out _); }
        bool AnyNumber(int r, int c) => cell(r, c).Kind == 1;
        bool Text(int r, int c) { var v = cell(r, c); return v.Kind == 2 || v.Kind == 1 && !Number(r, c); }   // a date shows as text
        bool Empty(int r, int c) => cell(r, c).Kind == 0;

        // a header row: the first row's value cells are text (or empty), and numbers follow below
        bool header = rows >= 2 && Enumerable.Range(cols >= 2 ? 1 : 0, cols >= 2 ? cols - 1 : 1).All(c => cell(0, c).Kind is 2 or 0)
                      && Enumerable.Range(cols >= 2 ? 1 : 0, cols >= 2 ? cols - 1 : 1).Any(c => cell(0, c).Kind == 2)
                      && Enumerable.Range(1, rows - 1).Any(r => Enumerable.Range(0, cols).Any(c => AnyNumber(r, c)));
        int top = header ? 1 : 0;
        // a categories column: text or dates down the first column, or an empty top-left cell over a header row
        bool catCol = cols >= 2 && ((header && Empty(0, 0)) || Enumerable.Range(top, rows - top).All(r => Empty(r, 0) || Text(r, 0)));
        bool scatterX = type == "scatter" && cols >= 2 && !catCol;
        int first = catCol || scatterX ? 1 : 0;
        int n = rows - top; bool truncated = false;
        if (n > MaxPoints) { n = MaxPoints; truncated = true; }
        var cats = new string[n];
        for (int i = 0; i < n; i++) cats[i] = catCol ? cell(top + i, 0).Text : (i + 1).ToString(CultureInfo.InvariantCulture);
        double?[]? x = null;
        if (scatterX) { x = new double?[n]; for (int i = 0; i < n; i++) { var v = cell(top + i, 0); x[i] = v.Kind == 1 ? v.Num : null; } }
        var series = new List<ChartSeries>();
        for (int c = first; c < cols; c++)
        {
            if (series.Count == MaxSeries) { truncated = true; break; }
            var vals = new double?[n];
            for (int i = 0; i < n; i++) { var v = cell(top + i, c); vals[i] = v.Kind == 1 ? v.Num : null; }
            string name = header ? cell(0, c).Text : (cols - first == 1 ? "Series" : "Series " + (series.Count + 1));
            series.Add(new ChartSeries(name.Length == 0 ? "Series " + (series.Count + 1) : name, vals));
        }
        return new(cats, series.ToArray(), x, truncated);
    }
}
