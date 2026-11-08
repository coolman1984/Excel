using System.Globalization;
using ExcelViewer.Core;

namespace ExcelApp.Data.Query;

public sealed record ResultPage(string ResultId, int From, string[][] Rows, int[][] Kinds);
public static class Format
{
    public static ResultPage Page(ResultStore store, string id, int from, int count)
    {
        if (from < 0 || count < 0 || count > 1000) throw new UserError("The requested page is invalid.");
        if (!store.TryGet(id, out var table, out var meta)) throw new UserError("That result is no longer available.");
        var columns = table.Columns.Where(c => !c.Name.StartsWith('_') || (meta.Kind == "diff" && c.Name == "_status")).ToArray();
        int length = Math.Min(count, Math.Max(0, table.Rows - from)); var rows = new string[length][]; var kinds = new int[length][];
        for (int i = 0; i < length; i++)
        {
            rows[i] = new string[columns.Length]; kinds[i] = new int[columns.Length]; int row = from + i;
            for (int j = 0; j < columns.Length; j++)
            {
                var c = columns[j]; if (c.IsNull(row)) { rows[i][j] = ""; continue; }
                kinds[i][j] = c.Type switch { ColType.Number or ColType.Date => 1, ColType.Text => 2, ColType.Bool => 3, _ => 0 };
                rows[i][j] = c.Type switch { ColType.Number => Number(c.Num(row)), ColType.Date => DateTime.FromOADate(c.Num(row)).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), _ => c.Text(row) };
            }
        }
        return new ResultPage(id, from, rows, kinds);
    }
    static string Number(double value)
    {
        if (double.IsFinite(value) && Math.Abs(value) < 1e15) value = Math.Round(value, 10, MidpointRounding.AwayFromZero);
        return value.ToString("G15", CultureInfo.InvariantCulture);
    }
}
