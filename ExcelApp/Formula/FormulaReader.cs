using System.Runtime.InteropServices;

namespace ExcelApp.Formula;

public static class FormulaReader
{
    public static IReadOnlyList<(int Sheet, int Row, int Col, string A1)> Read(dynamic workbook)
    {
        var result = new List<(int, int, int, string)>(); dynamic? sheets = null;
        try
        {
            sheets = workbook.Worksheets; int count = sheets.Count;
            for (int sheet = 1; sheet <= count; sheet++)
            {
                dynamic? ws = null, used = null, formulaCells = null, areas = null;
                try
                {
                    ws = sheets.Item[sheet]; used = ws.UsedRange;
                    try { formulaCells = used.SpecialCells(-4123); } catch (COMException) { continue; }
                    areas = formulaCells.Areas; int areaCount = areas.Count;
                    for (int i = 1; i <= areaCount; i++)
                    {
                        dynamic? area = null, areaRows = null, areaCols = null; try
                        {
                            area = areas.Item[i]; areaRows = area.Rows; areaCols = area.Columns; int row = area.Row, col = area.Column, rows = areaRows.Count, cols = areaCols.Count; object? formulas = area.Formula;
                            if (formulas is Array array) { int lo0 = array.GetLowerBound(0), lo1 = array.GetLowerBound(1); for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) if (array.GetValue(r + lo0, c + lo1) is string f && f.StartsWith('=')) result.Add((sheet - 1, row + r - 1, col + c - 1, f)); }
                            else if (formulas is string f && f.StartsWith('=')) result.Add((sheet - 1, row - 1, col - 1, f));
                        }
                        finally { Release(areaCols); Release(areaRows); Release(area); }
                    }
                }
                finally { Release(areas); Release(formulaCells); Release(used); Release(ws); }
            }
        }
        finally { Release(sheets); }
        return result;
    }
    static void Release(object? value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
}
