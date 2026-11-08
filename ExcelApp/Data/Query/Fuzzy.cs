using System.Text;

namespace ExcelApp.Data.Query;

public static class Fuzzy
{
    public static string Normalize(string value)
    {
        var s = new StringBuilder(value.Length); var upper = value.Trim().ToUpperInvariant();
        for (int i = 0; i < upper.Length;)
        {
            char c = upper[i]; if (char.IsWhiteSpace(c) || c is '-' or '_' or '.' or '/') { i++; continue; }
            if (char.IsDigit(c)) { int end = i; while (end < upper.Length && char.IsDigit(upper[end])) end++; int first = i; while (first < end - 1 && upper[first] == '0') first++; s.Append(upper, first, end - first); i = end; }
            else { s.Append(c); i++; }
        }
        return s.ToString();
    }
    public static double JaroWinkler(string a, string b)
    {
        if (a == b) return 1; if (a.Length == 0 || b.Length == 0) return 0;
        int window = Math.Max(0, Math.Max(a.Length, b.Length) / 2 - 1), matches = 0; var am = new bool[a.Length]; var bm = new bool[b.Length];
        for (int i = 0; i < a.Length; i++) for (int j = Math.Max(0, i - window); j < Math.Min(i + window + 1, b.Length); j++) if (!bm[j] && a[i] == b[j]) { am[i] = bm[j] = true; matches++; break; }
        if (matches == 0) return 0; int k = 0, transpositions = 0;
        for (int i = 0; i < a.Length; i++) if (am[i]) { while (!bm[k]) k++; if (a[i] != b[k++]) transpositions++; }
        double m = matches; double jaro = (m / a.Length + m / b.Length + (m - transpositions / 2.0) / m) / 3;
        int prefix = 0; while (prefix < Math.Min(4, Math.Min(a.Length, b.Length)) && a[prefix] == b[prefix]) prefix++;
        return jaro > 0.7 ? jaro + prefix * 0.1 * (1 - jaro) : jaro;
    }
}
