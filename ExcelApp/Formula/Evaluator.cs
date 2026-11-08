using System.Globalization;

namespace ExcelApp.Formula;

public sealed class Evaluator
{
    static readonly HashSet<string> Supported = new(StringComparer.Ordinal) { "SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "MIN", "MAX", "IF", "IFS", "IFERROR", "IFNA", "AND", "OR", "NOT", "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS", "INT", "MOD", "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "TEXT", "VALUE", "CONCAT", "CONCATENATE", "TEXTJOIN", "TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY", "EOMONTH", "EDATE", "WEEKDAY", "WEEKNUM", "NETWORKDAYS" };
    readonly IFormulaGrid grid; readonly int sheet;
    public Evaluator(IFormulaGrid grid, int sheet = 0) { this.grid = grid; this.sheet = sheet; }
    public FValue Evaluate(string formula) => Evaluate(new Parser(formula).Parse());
    static double? Number(FValue v) => v.Kind switch { ValueKind.Number or ValueKind.Bool => v.Number, ValueKind.Blank => 0, ValueKind.Text when double.TryParse(v.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) => n, _ => null };
    static string Text(FValue v) => v.Kind switch { ValueKind.Text or ValueKind.Error => v.Text ?? "", ValueKind.Bool => v.Number == 0 ? "FALSE" : "TRUE", ValueKind.Number => v.Number.ToString("G15", CultureInfo.InvariantCulture), _ => "" };
    int Sheet(Reference r) => r.Sheet == null ? sheet : grid.SheetIndex(r.Sheet);
    IEnumerable<FValue> Values(Node node)
    {
        if (node is Range range) { int s = Sheet(range.Start); if (s < 0) { yield return FValue.Error("#REF!"); yield break; } for (int r = Math.Min(range.Start.Row, range.End.Row); r <= Math.Max(range.Start.Row, range.End.Row); r++) for (int c = Math.Min(range.Start.Col, range.End.Col); c <= Math.Max(range.Start.Col, range.End.Col); c++) yield return grid.Get(s, r, c); }
        else yield return Evaluate(node);
    }
    public FValue Evaluate(Node node)
    {
        switch (node)
        {
            case Literal l: return l.Value;
            case Reference r: int s = Sheet(r); return s < 0 ? FValue.Error("#REF!") : grid.Get(s, r.Row, r.Col);
            case Range: return FValue.Error("#VALUE!");
            case Unary u: { var v = Evaluate(u.Value); if (v.Kind is ValueKind.Error or ValueKind.Unsupported) return v; var n = Number(v); return n == null ? FValue.Error("#VALUE!") : FValue.Num(u.Op == "-" ? -n.Value : n.Value / 100); }
            case Binary b:
            {
                var x = Evaluate(b.Left); if (x.Kind is ValueKind.Error or ValueKind.Unsupported) return x; var y = Evaluate(b.Right); if (y.Kind is ValueKind.Error or ValueKind.Unsupported) return y;
                if (b.Op == "&") return FValue.Str(Text(x) + Text(y)); var a = Number(x); var z = Number(y);
                if (b.Op is "=" or "<>" or "<" or ">" or "<=" or ">=") { int cmp = a.HasValue && z.HasValue ? a.Value.CompareTo(z.Value) : StringComparer.OrdinalIgnoreCase.Compare(Text(x), Text(y)); return FValue.Bool(b.Op switch { "=" => cmp == 0, "<>" => cmp != 0, "<" => cmp < 0, ">" => cmp > 0, "<=" => cmp <= 0, _ => cmp >= 0 }); }
                if (a == null || z == null) return FValue.Error("#VALUE!"); if (b.Op == "/" && z == 0) return FValue.Error("#DIV/0!"); double result = b.Op switch { "+" => a.Value + z.Value, "-" => a.Value - z.Value, "*" => a.Value * z.Value, "/" => a.Value / z.Value, "^" => Math.Pow(a.Value, z.Value), _ => double.NaN }; return double.IsFinite(result) ? FValue.Num(result) : FValue.Error("#NUM!");
            }
            case Call call: return Function(call);
            default: return FValue.Unsupported;
        }
    }
    static bool Criterion(FValue value, FValue criterion)
    {
        string expression = Text(criterion); string op = "="; foreach (string prefix in new[] { ">=", "<=", "<>", ">", "<", "=" }) if (expression.StartsWith(prefix, StringComparison.Ordinal)) { op = prefix; expression = expression[prefix.Length..]; break; }
        int cmp; bool match;
        if (double.TryParse(expression, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric) && Number(value) is double actual) { cmp = actual.CompareTo(numeric); match = cmp == 0; }
        else { string pattern = "^"; for (int i = 0; i < expression.Length; i++) { char ch = expression[i]; if (ch == '~' && i + 1 < expression.Length) pattern += System.Text.RegularExpressions.Regex.Escape(expression[++i].ToString()); else pattern += ch == '*' ? ".*" : ch == '?' ? "." : System.Text.RegularExpressions.Regex.Escape(ch.ToString()); } pattern += "$"; match = System.Text.RegularExpressions.Regex.IsMatch(Text(value), pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant); cmp = StringComparer.OrdinalIgnoreCase.Compare(Text(value), expression); }
        return op switch { "=" => match, "<>" => !match, ">" => cmp > 0, "<" => cmp < 0, ">=" => cmp >= 0, "<=" => cmp <= 0, _ => false };
    }
    FValue Conditional(Call call)
    {
        bool single = call.Name is "COUNTIF" or "SUMIF" or "AVERAGEIF"; bool count = call.Name.StartsWith("COUNT", StringComparison.Ordinal); bool avg = call.Name.StartsWith("AVERAGE", StringComparison.Ordinal);
        if (call.Args.Count < 2) return FValue.Error("#VALUE!"); Node? resultNode = count ? null : single ? (call.Args.Count >= 3 ? call.Args[2] : call.Args[0]) : call.Args[0];
        int begin = single || count ? 0 : 1; if ((call.Args.Count - begin) % 2 != 0 && !single) return FValue.Error("#VALUE!");
        var conditions = new List<(FValue[] range, FValue criterion)>(); int pairs = single ? 1 : (call.Args.Count - begin) / 2;
        for (int i = 0; i < pairs; i++) { int at = begin + 2 * i; conditions.Add((Values(call.Args[at]).ToArray(), Evaluate(call.Args[at + 1]))); }
        int length = conditions[0].range.Length; if (conditions.Any(c => c.range.Length != length)) return FValue.Error("#VALUE!"); var results = resultNode == null ? null : Values(resultNode).ToArray(); if (results != null && results.Length != length) return FValue.Error("#VALUE!"); double sum = 0; int hits = 0;
        for (int i = 0; i < length; i++) { if (!conditions.All(c => Criterion(c.range[i], c.criterion))) continue; if (count) hits++; else if (results![i].Kind == ValueKind.Number) { sum += results[i].Number; hits++; } }
        return count ? FValue.Num(hits) : avg ? hits == 0 ? FValue.Error("#DIV/0!") : FValue.Num(sum / hits) : FValue.Num(sum);
    }
    FValue Lookup(Call call)
    {
        var args = call.Args; string name = call.Name;
        if (name == "INDEX")
        {
            if (args.Count < 2 || args[0] is not Range range) return FValue.Error("#VALUE!"); var rowValue = Number(Evaluate(args[1])); var colValue = args.Count >= 3 ? Number(Evaluate(args[2])) : 1; if (rowValue == null || colValue == null) return FValue.Error("#VALUE!"); int row = (int)rowValue.Value, col = (int)colValue.Value, rows = Math.Abs(range.End.Row - range.Start.Row) + 1, cols = Math.Abs(range.End.Col - range.Start.Col) + 1; if (row < 1 || row > rows || col < 1 || col > cols) return FValue.Error("#REF!"); return grid.Get(Sheet(range.Start), Math.Min(range.Start.Row, range.End.Row) + row - 1, Math.Min(range.Start.Col, range.End.Col) + col - 1);
        }
        if (args.Count < 2) return FValue.Error("#VALUE!"); FValue target = Evaluate(args[0]); var search = name is "MATCH" or "XLOOKUP" ? Values(args[1]).ToArray() : Array.Empty<FValue>();
        if (name == "MATCH")
        {
            int mode = args.Count >= 3 && Number(Evaluate(args[2])) is double m ? (int)m : 1; int found = -1;
            for (int i = 0; i < search.Length; i++) { if (Criterion(search[i], target)) { found = i; break; } if (mode != 0 && Number(search[i]) is double n && Number(target) is double sought && (mode > 0 ? n <= sought : n >= sought)) found = i; }
            return found < 0 ? FValue.Error("#N/A") : FValue.Num(found + 1);
        }
        if (name == "XLOOKUP")
        {
            if (args.Count < 3) return FValue.Error("#VALUE!"); var returns = Values(args[2]).ToArray(); if (returns.Length != search.Length) return FValue.Error("#VALUE!");
            int mode = args.Count > 4 && Number(Evaluate(args[4])) is double mm ? (int)mm : 0, direction = args.Count > 5 && Number(Evaluate(args[5])) is double dd ? (int)dd : 1; if (mode is not (0 or -1 or 1 or 2) || direction is not (1 or -1)) return FValue.Error("#VALUE!");
            int found = -1; double best = double.PositiveInfinity; for (int n = 0; n < search.Length; n++) { int i = direction == 1 ? n : search.Length - 1 - n; var candidate = search[i]; bool equal = mode == 2 ? Criterion(candidate, target) : candidate.Kind == ValueKind.Text || target.Kind == ValueKind.Text ? string.Equals(Text(candidate), Text(target), StringComparison.OrdinalIgnoreCase) : Number(candidate) == Number(target); if (equal) { found = i; break; } if (mode is -1 or 1 && Number(candidate) is double x && Number(target) is double y && (mode == -1 ? x < y : x > y) && Math.Abs(x - y) < best) { best = Math.Abs(x - y); found = i; } }
            return found >= 0 ? returns[found] : args.Count >= 4 ? Evaluate(args[3]) : FValue.Error("#N/A");
        }
        if (args[1] is not Range matrix || args.Count < 3 || Number(Evaluate(args[2])) is not double offset) return FValue.Error("#VALUE!");
        int startRow = Math.Min(matrix.Start.Row, matrix.End.Row), startCol = Math.Min(matrix.Start.Col, matrix.End.Col), rowCount = Math.Abs(matrix.End.Row - matrix.Start.Row) + 1, colCount = Math.Abs(matrix.End.Col - matrix.Start.Col) + 1, position = (int)offset - 1; bool horizontal = name == "HLOOKUP"; if (position < 0 || position >= (horizontal ? rowCount : colCount)) return FValue.Error("#REF!");
        bool approximate = args.Count < 4 || Number(Evaluate(args[3])) != 0; int match = -1; for (int i = 0; i < (horizontal ? colCount : rowCount); i++) { var candidate = grid.Get(Sheet(matrix.Start), startRow + (horizontal ? 0 : i), startCol + (horizontal ? i : 0)); if (Criterion(candidate, target)) { match = i; break; } if (approximate && Number(candidate) is double v && Number(target) is double t && v <= t) match = i; }
        return match < 0 ? FValue.Error("#N/A") : grid.Get(Sheet(matrix.Start), startRow + (horizontal ? position : match), startCol + (horizontal ? match : position));
    }
    FValue Function(Call call)
    {
        var args = call.Args; string name = call.Name;
        if (!Supported.Contains(name)) return FValue.Unsupported;
        if (name is "INDEX" or "MATCH" or "XLOOKUP" or "VLOOKUP" or "HLOOKUP") return Lookup(call);
        if (name is "COUNTIF" or "COUNTIFS" or "SUMIF" or "SUMIFS" or "AVERAGEIF" or "AVERAGEIFS") return Conditional(call);
        if (name == "IFS")
        {
            if (args.Count == 0 || args.Count % 2 != 0) return FValue.Error("#VALUE!");
            for (int i = 0; i < args.Count; i += 2) { var condition = Evaluate(args[i]); if (condition.Kind is ValueKind.Error or ValueKind.Unsupported) return condition; var truth = Number(condition); if (truth == null) return FValue.Error("#VALUE!"); if (truth != 0) return Evaluate(args[i + 1]); } return FValue.Error("#N/A");
        }
        if (name is "IF" or "IFERROR" or "IFNA")
        {
            if (args.Count < 2 || args.Count > 3) return FValue.Error("#VALUE!"); var value = Evaluate(args[0]);
            if (name == "IF") { var truth = Number(value); if (truth == null) return value.Kind == ValueKind.Error ? value : FValue.Error("#VALUE!"); return truth != 0 ? Evaluate(args[1]) : args.Count == 3 ? Evaluate(args[2]) : FValue.Bool(false); }
            return value.Kind == ValueKind.Error && (name == "IFERROR" || value.Text == "#N/A") ? Evaluate(args[1]) : value;
        }
        if (name == "TODAY" && args.Count == 0) return FValue.Num(DateTime.Today.ToOADate());
        if (name == "NOW" && args.Count == 0) return FValue.Num(DateTime.Now.ToOADate());
        var values = args.SelectMany(Values).ToArray(); var error = values.FirstOrDefault(v => v.Kind is ValueKind.Error or ValueKind.Unsupported); if (error.Kind is ValueKind.Error or ValueKind.Unsupported) return error;
        var nums = values.Where(v => v.Kind == ValueKind.Number).Select(v => v.Number).ToArray();
        if (name == "SUM") return FValue.Num(nums.Sum()); if (name == "COUNT") return FValue.Num(nums.Length); if (name == "COUNTA") return FValue.Num(values.Count(v => v.Kind != ValueKind.Blank)); if (name == "COUNTBLANK") return FValue.Num(values.Count(v => v.Kind == ValueKind.Blank || (v.Kind == ValueKind.Text && v.Text == "")));
        if (name is "AVERAGE" or "MIN" or "MAX") { if (nums.Length == 0) return name == "AVERAGE" ? FValue.Error("#DIV/0!") : FValue.Num(0); return FValue.Num(name == "AVERAGE" ? nums.Average() : name == "MIN" ? nums.Min() : nums.Max()); }
        if (name is "AND" or "OR") { var logical = values.Select(Number).Where(v => v.HasValue).Select(v => v!.Value != 0).ToArray(); return FValue.Bool(name == "AND" ? logical.All(x => x) : logical.Any(x => x)); }
        if (name == "CONCAT" || name == "CONCATENATE") return FValue.Str(string.Concat(values.Select(Text)));
        if (values.Length == 0) return FValue.Unsupported;
        var n0 = Number(values[0]);
        if (name == "NOT") return n0.HasValue ? FValue.Bool(n0 == 0) : FValue.Error("#VALUE!");
        if (name is "LEN" or "TRIM" or "UPPER" or "LOWER") { string s = Text(values[0]); return name switch { "LEN" => FValue.Num(s.Length), "TRIM" => FValue.Str(string.Join(" ", s.Split(' ', StringSplitOptions.RemoveEmptyEntries))), "UPPER" => FValue.Str(s.ToUpperInvariant()), _ => FValue.Str(s.ToLowerInvariant()) }; }
        if (name is "LEFT" or "RIGHT" or "MID")
        {
            string text = Text(values[0]); int start = 0, count;
            if (name == "MID") { if (values.Length < 3 || Number(values[1]) is not double st || Number(values[2]) is not double len || st < 1 || len < 0) return FValue.Error("#VALUE!"); start = (int)st - 1; count = (int)len; }
            else { count = values.Length < 2 ? 1 : Number(values[1]) is double v ? (int)v : -1; if (count < 0) return FValue.Error("#VALUE!"); if (name == "RIGHT") start = Math.Max(0, text.Length - count); }
            if (start >= text.Length) return FValue.Str(""); return FValue.Str(text.Substring(start, Math.Min(count, text.Length - start)));
        }
        if (name == "TEXTJOIN" && values.Length >= 2) { string delimiter = Text(values[0]); bool skip = Number(values[1]) != 0; return FValue.Str(string.Join(delimiter, values.Skip(2).Select(Text).Where(x => !skip || x.Length > 0))); }
        if (name == "TEXT" && values.Length == 2 && n0.HasValue) { var shown = new ExcelViewer.Core.NumFmt(Text(values[1])).Number(n0.Value); return shown.HasValue ? FValue.Str(shown.Value.Text) : FValue.Unsupported; }
        if (name == "DATE" && values.Length == 3 && values.All(v => Number(v).HasValue))
        {
            try { int year = (int)Number(values[0])!.Value; if (year >= 0 && year <= 1899) year += 1900; int month = (int)Number(values[1])!.Value, day = (int)Number(values[2])!.Value; return FValue.Num(new DateTime(year, 1, 1).AddMonths(month - 1).AddDays(day - 1).ToOADate()); } catch (ArgumentOutOfRangeException) { return FValue.Error("#NUM!"); }
        }
        if (name is "YEAR" or "MONTH" or "DAY" or "EDATE" or "EOMONTH" or "WEEKDAY" or "WEEKNUM")
        {
            if (n0 == null) return FValue.Error("#VALUE!"); try
            {
                var date = DateTime.FromOADate(n0.Value); if (name == "YEAR") return FValue.Num(date.Year); if (name == "MONTH") return FValue.Num(date.Month); if (name == "DAY") return FValue.Num(date.Day);
                if (name is "EDATE" or "EOMONTH") { if (values.Length < 2 || Number(values[1]) is not double months) return FValue.Error("#VALUE!"); var shifted = date.AddMonths((int)months); return FValue.Num((name == "EOMONTH" ? new DateTime(shifted.Year, shifted.Month, DateTime.DaysInMonth(shifted.Year, shifted.Month)) : shifted).ToOADate()); }
                int mode = values.Length > 1 && Number(values[1]) is double m ? (int)m : 1;
                if (name == "WEEKDAY") { int dow = (int)date.DayOfWeek; return mode switch { 1 => FValue.Num(dow + 1), 2 => FValue.Num((dow + 6) % 7 + 1), 3 => FValue.Num((dow + 6) % 7), _ => FValue.Error("#NUM!") }; }
                if (mode is not (1 or 2 or 21)) return FValue.Error("#NUM!"); return FValue.Num(mode == 21 ? ISOWeek.GetWeekOfYear(date) : CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay, mode == 1 ? DayOfWeek.Sunday : DayOfWeek.Monday));
            } catch (ArgumentException) { return FValue.Error("#NUM!"); }
        }
        if (name == "NETWORKDAYS" && values.Length >= 2 && n0.HasValue && Number(values[1]) is double last)
        {
            try { var start = DateTime.FromOADate(n0.Value).Date; var end = DateTime.FromOADate(last).Date; int sign = start <= end ? 1 : -1; if (sign < 0) (start, end) = (end, start); var holidays = values.Skip(2).Where(v => v.Kind == ValueKind.Number).Select(v => DateTime.FromOADate(v.Number).Date).ToHashSet(); int count = 0; for (var d = start; d <= end; d = d.AddDays(1)) if (d.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday) && !holidays.Contains(d)) count++; return FValue.Num(sign * count); } catch (ArgumentException) { return FValue.Error("#NUM!"); }
        }
        if (name is "ABS" or "INT") return n0.HasValue ? FValue.Num(name == "ABS" ? Math.Abs(n0.Value) : Math.Floor(n0.Value)) : FValue.Error("#VALUE!");
        if (name == "VALUE") return n0.HasValue ? FValue.Num(n0.Value) : FValue.Error("#VALUE!");
        if (values.Length >= 2 && name == "MOD") { var n1 = Number(values[1]); return n0 == null || n1 == null ? FValue.Error("#VALUE!") : n1 == 0 ? FValue.Error("#DIV/0!") : FValue.Num(n0.Value - n1.Value * Math.Floor(n0.Value / n1.Value)); }
        if (values.Length >= 2 && name is "ROUND" or "ROUNDUP" or "ROUNDDOWN") { var digits = Number(values[1]); if (n0 == null || digits == null || Math.Abs(digits.Value) > 15) return FValue.Error("#NUM!"); double scale = Math.Pow(10, Math.Truncate(digits.Value)); double scaled = n0.Value * scale; double rounded = name == "ROUND" ? Math.Round(scaled, MidpointRounding.AwayFromZero) : name == "ROUNDUP" ? Math.Sign(scaled) * Math.Ceiling(Math.Abs(scaled)) : Math.Truncate(scaled); return FValue.Num(rounded / scale); }
        return FValue.Unsupported;
    }
}
