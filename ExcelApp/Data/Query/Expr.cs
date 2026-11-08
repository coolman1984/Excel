using System.Globalization;

namespace ExcelApp.Data.Query;

/// <summary>Parsed numerical expression. Nulls propagate except through coalesce.</summary>
public sealed class Expr
{
    readonly Func<Table, int, double?> evaluate;
    Expr(Func<Table, int, double?> evaluate) => this.evaluate = evaluate;
    public double? Evaluate(Table table, int row) => evaluate(table, row);
    public static Expr Parse(string text) => new(new Parser(text).Parse());

    sealed class Parser
    {
        readonly string text; int pos;
        public Parser(string text) => this.text = text;
        void Space() { while (pos < text.Length && char.IsWhiteSpace(text[pos])) pos++; }
        bool Take(char c) { Space(); if (pos < text.Length && text[pos] == c) { pos++; return true; } return false; }
        Exception Invalid() => new FormatException("Invalid expression.");
        public Func<Table, int, double?> Parse() { var e = Sum(); Space(); if (pos != text.Length) throw Invalid(); return e; }
        Func<Table, int, double?> Sum()
        {
            var left = Product();
            while (true) { char op; if (Take('+')) op = '+'; else if (Take('-')) op = '-'; else break; var a = left; var b = Product(); left = (t, r) => op == '+' ? a(t, r) + b(t, r) : a(t, r) - b(t, r); }
            return left;
        }
        Func<Table, int, double?> Product()
        {
            var left = Primary();
            while (true) { char op; if (Take('*')) op = '*'; else if (Take('/')) op = '/'; else break; var a = left; var b = Primary(); left = (t, r) => { var x = a(t, r); var y = b(t, r); return op == '*' ? x * y : y == 0 ? null : x / y; }; }
            return left;
        }
        Func<Table, int, double?> Primary()
        {
            if (Take('-')) { var x = Primary(); return (t, r) => -x(t, r); }
            if (Take('+')) return Primary();
            if (Take('(')) { var x = Sum(); if (!Take(')')) throw Invalid(); return x; }
            Space(); int start = pos;
            if (pos < text.Length && (char.IsDigit(text[pos]) || text[pos] == '.'))
            {
                while (pos < text.Length && (char.IsDigit(text[pos]) || text[pos] == '.')) pos++;
                if (!double.TryParse(text.AsSpan(start, pos - start), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var n)) throw Invalid();
                return (_, _) => n;
            }
            string name;
            if (Take('[')) { start = pos; while (pos < text.Length && text[pos] != ']') pos++; if (pos == text.Length || pos == start) throw Invalid(); name = text[start..pos++]; }
            else { Space(); start = pos; while (pos < text.Length && (char.IsLetterOrDigit(text[pos]) || text[pos] == '_')) pos++; if (start == pos) throw Invalid(); name = text[start..pos]; }
            if (Take('('))
            {
                var args = new List<Func<Table, int, double?>>();
                if (!Take(')')) { do { args.Add(Sum()); } while (Take(',')); if (!Take(')')) throw Invalid(); }
                var function = name.ToLowerInvariant();
                if (((function == "abs" || function == "month") && args.Count != 1) || ((function == "round" || function == "coalesce" || function == "pct") && args.Count != 2) || function is not ("abs" or "month" or "round" or "coalesce" or "pct")) throw Invalid();
                return (t, r) => { var a = args[0](t, r); if (function == "abs") return a.HasValue ? Math.Abs(a.Value) : null; if (function == "month") { if (!a.HasValue) return null; try { var d = DateTime.FromOADate(a.Value); return new DateTime(d.Year, d.Month, 1).ToOADate(); } catch (ArgumentException) { return null; } } var b = args[1](t, r); if (function == "coalesce") return a ?? b; if (!a.HasValue || !b.HasValue) return null; if (function == "pct") return b == 0 ? null : (a - b) / b * 100; if (b < 0 || b > 15 || b != Math.Truncate(b.Value)) return null; return Math.Round(a.Value, (int)b.Value, MidpointRounding.AwayFromZero); };
            }
            return (t, r) => { var c = t[name] ?? throw new ArgumentException("Expression refers to an unknown column."); if (c.IsNull(r) || c.Type == ColType.Text) return null; return c.Num(r); };
        }
    }
}
