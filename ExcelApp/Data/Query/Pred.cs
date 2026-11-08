using System.Globalization;
using System.Text.Json;

namespace ExcelApp.Data.Query;

/// <summary>Compiles the wire predicate against a table once, including duplicate-key indexes.</summary>
public static class Pred
{
    public static Func<int, bool> Compile(JsonElement predicate, Table table, DateTime? now = null)
    {
        var clock = now ?? DateTime.Today;
        Column Col(string name) => table[name] ?? throw new ArgumentException("Predicate refers to an unknown column.");
        string Str(JsonElement p, string name) => p.GetProperty(name).GetString() ?? throw new FormatException("Invalid predicate.");
        object? Literal(JsonElement e, Column c)
        {
            if (e.ValueKind == JsonValueKind.Number) return e.GetDouble();
            if (e.ValueKind is JsonValueKind.True or JsonValueKind.False) return e.GetBoolean() ? 1d : 0d;
            if (e.ValueKind != JsonValueKind.String) throw new FormatException("Invalid predicate value.");
            var s = e.GetString()!;
            if (c.Type == ColType.Date && RelativeDate.TryResolve(s, clock, out var date)) return date.ToOADate();
            if (c.Type == ColType.Date && DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) return date.ToOADate();
            return s;
        }
        static int Compare(Column c, int row, object? value)
        {
            if (value is double n && c.Type != ColType.Text) return c.Num(row).CompareTo(n);
            return StringComparer.OrdinalIgnoreCase.Compare(c.Text(row), Convert.ToString(value, CultureInfo.InvariantCulture));
        }
        static bool Match(string op, int cmp) => op switch { "eq" => cmp == 0, "ne" => cmp != 0, "gt" => cmp > 0, "ge" => cmp >= 0, "lt" => cmp < 0, "le" => cmp <= 0, _ => throw new FormatException("Invalid comparison.") };
        Func<int, bool> Build(JsonElement p)
        {
            var op = Str(p, "op");
            if (op is "and" or "or") { var items = p.GetProperty("items").EnumerateArray().Select(Build).ToArray(); return r => op == "and" ? items.All(f => f(r)) : items.Any(f => f(r)); }
            if (op == "not") { var f = Build(p.GetProperty("item")); return r => !f(r); }
            if (op == "duplicate")
            {
                var cols = p.GetProperty("cols").EnumerateArray().Select(x => Col(x.GetString()!)).ToArray();
                if (cols.Length == 0) throw new FormatException("Duplicate requires columns.");
                string Key(int r) => string.Concat(cols.Select(c => c.IsNull(r) ? "-1:" : c.Text(r).Length.ToString(CultureInfo.InvariantCulture) + ":" + c.Text(r).ToUpperInvariant()));
                var counts = new Dictionary<string, int>(StringComparer.Ordinal);
                for (int r = 0; r < table.Rows; r++) { var key = Key(r); counts.TryGetValue(key, out int n); counts[key] = n + 1; }
                return r => counts[Key(r)] > 1;
            }
            var col = Col(Str(p, "col"));
            if (op == "isNull") return r => col.IsNull(r);
            if (op is "blank" or "notBlank") return r => (col.IsNull(r) || col.Text(r).Length == 0) == (op == "blank");
            if (op == "colCompare")
            {
                var other = Col(Str(p, "other")); var cmp = Str(p, "cmp");
                double tolerance = p.TryGetProperty("tolerancePct", out var tol) ? tol.GetDouble() : 0;
                return r => { if (col.IsNull(r) || other.IsNull(r)) return false; int comparison;
                    if (col.Type != ColType.Text && other.Type != ColType.Text) { var a = col.Num(r); var b = other.Num(r); comparison = Math.Abs(a - b) <= Math.Abs(b) * tolerance / 100 ? 0 : a.CompareTo(b); }
                    else comparison = StringComparer.OrdinalIgnoreCase.Compare(col.Text(r), other.Text(r));
                    return Match(cmp, comparison); };
            }
            if (op == "between") { var a = Literal(p.GetProperty("a"), col); var b = Literal(p.GetProperty("b"), col); return r => !col.IsNull(r) && Compare(col, r, a) >= 0 && Compare(col, r, b) <= 0; }
            if (op == "in") { var values = p.GetProperty("values").EnumerateArray().Select(x => Literal(x, col)).ToArray(); return r => !col.IsNull(r) && values.Any(v => Compare(col, r, v) == 0); }
            var value = Literal(p.GetProperty("value"), col);
            if (op is "contains" or "begins" or "ends") { var s = Convert.ToString(value, CultureInfo.InvariantCulture)!; return r => !col.IsNull(r) && (op == "contains" ? col.Text(r).Contains(s, StringComparison.OrdinalIgnoreCase) : op == "begins" ? col.Text(r).StartsWith(s, StringComparison.OrdinalIgnoreCase) : col.Text(r).EndsWith(s, StringComparison.OrdinalIgnoreCase)); }
            if (op is not ("eq" or "ne" or "gt" or "ge" or "lt" or "le")) throw new FormatException("Invalid predicate operation.");
            return r => !col.IsNull(r) && Match(op, Compare(col, r, value));
        }
        return Build(predicate);
    }
}
