// Excel addresses for AI agents: "B7", "A1:D20", "C:C" and "3:3", 1-based as Excel shows them.
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelApp.Agent;

public static class A1
{
    static readonly Regex Cell = new(@"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$", RegexOptions.CultureInvariant);
    static readonly Regex ColOnly = new(@"^\$?([A-Za-z]{1,3})$", RegexOptions.CultureInvariant);
    static readonly Regex RowOnly = new(@"^\$?(\d{1,7})$", RegexOptions.CultureInvariant);
    public const int MaxRow = 1_048_576, MaxCol = 16_384;

    public static int ColNumber(string letters) { int n = 0; foreach (char c in letters.ToUpperInvariant()) n = n * 26 + c - 'A' + 1; return n; }
    public static string ColLetters(int n) { var s = ""; while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; } return s; }

    /// <summary>First row, first column, last row, last column (1-based). A whole column or row stops at <paramref name="rows"/> / <paramref name="cols"/>.</summary>
    public static (int R1, int C1, int R2, int C2) Parse(string text, int rows = MaxRow, int cols = MaxCol)
    {
        var parts = (text ?? "").Trim().Split(':');
        if (parts.Length is < 1 or > 2 || parts[0].Length == 0) throw Bad(text);
        (int r, int c, bool wholeCol, bool wholeRow) One(string p)
        {
            var m = Cell.Match(p);
            if (m.Success) return (int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture), ColNumber(m.Groups[1].Value), false, false);
            m = ColOnly.Match(p); if (m.Success) return (0, ColNumber(m.Groups[1].Value), true, false);
            m = RowOnly.Match(p); if (m.Success) return (int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture), 0, false, true);
            throw Bad(text);
        }
        var a = One(parts[0]); var b = parts.Length == 2 ? One(parts[1]) : a;
        if (a.wholeCol != b.wholeCol || a.wholeRow != b.wholeRow) throw Bad(text);
        if ((a.wholeCol || a.wholeRow) && parts.Length == 1) throw Bad(text);
        int r1 = a.wholeCol ? 1 : a.r, r2 = b.wholeCol ? Math.Max(1, rows) : b.r, c1 = a.wholeRow ? 1 : a.c, c2 = b.wholeRow ? Math.Max(1, cols) : b.c;
        if (r1 > r2) (r1, r2) = (r2, r1); if (c1 > c2) (c1, c2) = (c2, c1);
        if (r1 < 1 || r2 > MaxRow || c1 < 1 || c2 > MaxCol) throw Bad(text);
        return (r1, c1, r2, c2);
    }
    public static string Format(int r1, int c1, int r2, int c2) => r1 == r2 && c1 == c2 ? ColLetters(c1) + r1 : ColLetters(c1) + r1 + ":" + ColLetters(c2) + r2;
    static UserError Bad(string? text) => new($"\"{text}\" is not a cell or range. Write it as Excel does, such as B7, A1:D20, C:C or 3:5.");
}
