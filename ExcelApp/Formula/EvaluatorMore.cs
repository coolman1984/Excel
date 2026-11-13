// More of Excel's worksheet functions: maths, statistics, text, information, dates and finance.
using System.Globalization;
using System.Text;

namespace ExcelApp.Formula;

public sealed partial class Evaluator
{
    static readonly HashSet<string> Extra = new(StringComparer.Ordinal)
    {
        "PRODUCT", "SUMPRODUCT", "SUMSQ", "POWER", "SQRT", "EXP", "LN", "LOG", "LOG10", "PI", "SIGN", "TRUNC", "CEILING", "CEILING.MATH", "FLOOR", "FLOOR.MATH",
        "MROUND", "EVEN", "ODD", "FACT", "GCD", "LCM", "RAND", "RANDBETWEEN", "QUOTIENT", "RADIANS", "DEGREES", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "ATAN2",
        "SUBTOTAL", "MEDIAN", "MODE", "MODE.SNGL", "STDEV", "STDEV.S", "STDEV.P", "STDEVP", "VAR", "VAR.S", "VAR.P", "VARP", "LARGE", "SMALL", "RANK", "RANK.EQ",
        "MAXIFS", "MINIFS", "PERCENTILE", "PERCENTILE.INC", "QUARTILE", "QUARTILE.INC", "AVERAGEA", "MAXA", "MINA", "COUNTUNIQUE",
        "FIND", "SEARCH", "SUBSTITUTE", "REPLACE", "REPT", "EXACT", "PROPER", "CHAR", "CODE", "UNICHAR", "UNICODE", "CLEAN", "T", "N", "FIXED", "TEXTBEFORE", "TEXTAFTER", "NUMBERVALUE",
        "XOR", "TRUE", "FALSE", "ISBLANK", "ISNUMBER", "ISTEXT", "ISNONTEXT", "ISLOGICAL", "ISERROR", "ISERR", "ISNA", "ISEVEN", "ISODD", "NA", "CHOOSE", "SWITCH",
        "ROW", "COLUMN", "ROWS", "COLUMNS",
        "HOUR", "MINUTE", "SECOND", "TIME", "DATEDIF", "DAYS", "DATEVALUE", "WORKDAY", "YEARFRAC",
        "PMT", "PV", "FV", "NPV", "NPER", "IPMT", "PPMT",
    };

    /// <summary>Every function name the engine calculates.</summary>
    public static IReadOnlyCollection<string> Functions => Supported.Concat(Extra).OrderBy(x => x, StringComparer.Ordinal).ToList();

    static readonly Random Rng = new();
    static bool IsErr(FValue v) => v.Kind is ValueKind.Error or ValueKind.Unsupported;

    (int R1, int C1, int R2, int C2) Bounds(Range range, int s)
    {
        var (rows, cols) = s >= 0 ? grid.Extent(s) : (1048576, 16384);
        int r1 = Math.Min(range.Start.Row, range.End.Row), r2 = Math.Min(Math.Max(range.Start.Row, range.End.Row), Math.Max(r1, rows - 1));
        int c1 = Math.Min(range.Start.Col, range.End.Col), c2 = Math.Min(Math.Max(range.Start.Col, range.End.Col), Math.Max(c1, cols - 1));
        return (r1, c1, r2, c2);
    }

    FValue? More(Call call)
    {
        string name = call.Name; var args = call.Args;
        if (!Extra.Contains(name)) return null;
        FValue Arg(int i) => i < args.Count ? Evaluate(args[i]) : FValue.Blank;
        double? NumArg(int i, out FValue err) { var v = Arg(i); err = v; if (IsErr(v)) return null; var n = Number(v); if (n == null) err = FValue.Error("#VALUE!"); return n; }

        // ---- information: these see errors instead of passing them on
        switch (name)
        {
            case "ISBLANK": return FValue.Bool(Arg(0).Kind == ValueKind.Blank);
            case "ISNUMBER": return FValue.Bool(Arg(0).Kind == ValueKind.Number);
            case "ISTEXT": return FValue.Bool(Arg(0).Kind == ValueKind.Text);
            case "ISNONTEXT": return FValue.Bool(Arg(0).Kind != ValueKind.Text);
            case "ISLOGICAL": return FValue.Bool(Arg(0).Kind == ValueKind.Bool);
            case "ISERROR": return FValue.Bool(IsErr(Arg(0)));
            case "ISERR": { var v = Arg(0); return FValue.Bool(IsErr(v) && v.Text != "#N/A"); }
            case "ISNA": return FValue.Bool(Arg(0) is { Kind: ValueKind.Error, Text: "#N/A" });
            case "NA": return FValue.Error("#N/A");
            case "TRUE": return FValue.Bool(true);
            case "FALSE": return FValue.Bool(false);
            case "PI": return FValue.Num(Math.PI);
            case "RAND": return FValue.Num(Rng.NextDouble());
            case "ROW":
                if (args.Count == 0) return Row >= 0 ? FValue.Num(Row + 1) : FValue.Error("#VALUE!");
                return args[0] switch { Reference r => FValue.Num(r.Row + 1), Range g => FValue.Num(Math.Min(g.Start.Row, g.End.Row) + 1), _ => FValue.Error("#VALUE!") };
            case "COLUMN":
                if (args.Count == 0) return Col >= 0 ? FValue.Num(Col + 1) : FValue.Error("#VALUE!");
                return args[0] switch { Reference r => FValue.Num(r.Col + 1), Range g => FValue.Num(Math.Min(g.Start.Col, g.End.Col) + 1), _ => FValue.Error("#VALUE!") };
            case "ROWS":
            case "COLUMNS":
                if (args.Count != 1) return FValue.Error("#VALUE!");
                if (args[0] is Reference) return FValue.Num(1);
                if (args[0] is Range rg) return FValue.Num(name == "ROWS" ? Math.Abs(rg.End.Row - rg.Start.Row) + 1 : Math.Abs(rg.End.Col - rg.Start.Col) + 1);
                return FValue.Error("#VALUE!");
            case "CHOOSE":
            {
                if (args.Count < 2) return FValue.Error("#VALUE!");
                var idx = NumArg(0, out var e); if (idx == null) return e;
                int k = (int)Math.Truncate(idx.Value); return k < 1 || k >= args.Count ? FValue.Error("#VALUE!") : Evaluate(args[k]);
            }
            case "SWITCH":
            {
                if (args.Count < 3) return FValue.Error("#VALUE!");
                var target = Arg(0); if (IsErr(target)) return target;
                int i = 1;
                for (; i + 1 < args.Count; i += 2) { var cand = Evaluate(args[i]); if (IsErr(cand)) return cand; if (Same(target, cand)) return Evaluate(args[i + 1]); }
                return i < args.Count ? Evaluate(args[i]) : FValue.Error("#N/A");
            }
            case "SUMPRODUCT":
            {
                if (args.Count == 0) return FValue.Error("#VALUE!");
                var arrays = args.Select(a => Values(a).ToArray()).ToList();
                int n = arrays[0].Length; if (arrays.Any(a => a.Length != n)) return FValue.Error("#VALUE!");
                double sum = 0;
                for (int i = 0; i < n; i++)
                {
                    double p = 1;
                    foreach (var a in arrays) { if (IsErr(a[i])) return a[i]; p *= a[i].Kind == ValueKind.Number ? a[i].Number : 0; }
                    sum += p;
                }
                return FValue.Num(sum);
            }
            case "SUBTOTAL":
            {
                if (args.Count < 2) return FValue.Error("#VALUE!");
                var f = NumArg(0, out var e); if (f == null) return e;
                int code = (int)f.Value % 100;
                string? fn = code switch { 1 => "AVERAGE", 2 => "COUNT", 3 => "COUNTA", 4 => "MAX", 5 => "MIN", 6 => "PRODUCT", 7 => "STDEV", 8 => "STDEVP", 9 => "SUM", 10 => "VAR", 11 => "VARP", _ => null };
                return fn == null ? FValue.Error("#VALUE!") : Evaluate(new Call(fn, args.Skip(1).ToList()));
            }
            case "MAXIFS":
            case "MINIFS":
            {
                if (args.Count < 3 || (args.Count - 1) % 2 != 0) return FValue.Error("#VALUE!");
                var values = Values(args[0]).ToArray(); var conds = new List<(FValue[] range, FValue crit)>();
                for (int i = 1; i < args.Count; i += 2) conds.Add((Values(args[i]).ToArray(), Evaluate(args[i + 1])));
                if (conds.Any(c => c.range.Length != values.Length)) return FValue.Error("#VALUE!");
                double? best = null;
                for (int i = 0; i < values.Length; i++)
                {
                    if (values[i].Kind != ValueKind.Number || !conds.All(c => Criterion(c.range[i], c.crit))) continue;
                    best = best == null ? values[i].Number : name == "MAXIFS" ? Math.Max(best.Value, values[i].Number) : Math.Min(best.Value, values[i].Number);
                }
                return FValue.Num(best ?? 0);
            }
            case "RANK":
            case "RANK.EQ":
            {
                if (args.Count < 2) return FValue.Error("#VALUE!");
                var x = NumArg(0, out var e); if (x == null) return e;
                var list = Values(args[1]).Where(v => v.Kind == ValueKind.Number).Select(v => v.Number).ToList();
                bool asc = args.Count > 2 && Number(Arg(2)) is double o && o != 0;
                if (!list.Contains(x.Value)) return FValue.Error("#N/A");
                return FValue.Num(1 + list.Count(v => asc ? v < x.Value : v > x.Value));
            }
            case "LARGE":
            case "SMALL":
            case "PERCENTILE":
            case "PERCENTILE.INC":
            case "QUARTILE":
            case "QUARTILE.INC":
            {
                if (args.Count < 2) return FValue.Error("#VALUE!");
                var data = Values(args[0]).ToArray(); var bad = data.FirstOrDefault(IsErr); if (IsErr(bad)) return bad;
                var nums = data.Where(v => v.Kind == ValueKind.Number).Select(v => v.Number).OrderBy(v => v).ToArray();
                var k = NumArg(1, out var e); if (k == null) return e;
                if (nums.Length == 0) return FValue.Error("#NUM!");
                if (name is "LARGE" or "SMALL")
                {
                    int i = (int)Math.Ceiling(k.Value); if (i < 1 || i > nums.Length) return FValue.Error("#NUM!");
                    return FValue.Num(name == "SMALL" ? nums[i - 1] : nums[nums.Length - i]);
                }
                double p = name.StartsWith("QUARTILE", StringComparison.Ordinal) ? Math.Truncate(k.Value) / 4 : k.Value;
                if (p < 0 || p > 1) return FValue.Error("#NUM!");
                double pos = p * (nums.Length - 1); int lo = (int)Math.Floor(pos); double frac = pos - lo;
                return FValue.Num(lo + 1 < nums.Length ? nums[lo] + frac * (nums[lo + 1] - nums[lo]) : nums[lo]);
            }
            case "AVERAGEA":
            case "MAXA":
            case "MINA":
            {
                var all = args.SelectMany(Values).ToArray(); var bad = all.FirstOrDefault(IsErr); if (IsErr(bad)) return bad;
                var nums = all.Where(v => v.Kind != ValueKind.Blank).Select(v => v.Kind == ValueKind.Text ? 0 : v.Number).ToArray();
                if (nums.Length == 0) return name == "AVERAGEA" ? FValue.Error("#DIV/0!") : FValue.Num(0);
                return FValue.Num(name == "AVERAGEA" ? nums.Average() : name == "MAXA" ? nums.Max() : nums.Min());
            }
            case "COUNTUNIQUE":
                return FValue.Num(args.SelectMany(Values).Where(v => v.Kind != ValueKind.Blank).Select(v => v.Kind + ":" + (v.Kind == ValueKind.Number ? v.Number.ToString("R", CultureInfo.InvariantCulture) : (v.Text ?? "").ToUpperInvariant())).Distinct().Count());
        }

        // ---- the rest take plain values and pass errors on
        var vals = args.SelectMany(Values).ToArray();
        var error = vals.FirstOrDefault(IsErr); if (IsErr(error)) return error;
        var numbers = vals.Where(v => v.Kind == ValueKind.Number).Select(v => v.Number).ToArray();
        double? N(int i) => i < vals.Length ? Number(vals[i]) : null;
        string S(int i) => i < vals.Length ? Text(vals[i]) : "";
        FValue Num(double v) => double.IsFinite(v) ? FValue.Num(v) : FValue.Error("#NUM!");
        FValue Need1(Func<double, FValue> f) => N(0) is double a ? f(a) : FValue.Error("#VALUE!");
        FValue Need2(Func<double, double, FValue> f) => N(0) is double a && N(1) is double b ? f(a, b) : FValue.Error("#VALUE!");

        switch (name)
        {
            // maths
            case "PRODUCT": return Num(numbers.Aggregate(numbers.Length == 0 ? 0 : 1.0, (a, b) => a * b));
            case "SUMSQ": return Num(numbers.Sum(x => x * x));
            case "POWER": return Need2((a, b) => Num(Math.Pow(a, b)));
            case "SQRT": return Need1(a => a < 0 ? FValue.Error("#NUM!") : Num(Math.Sqrt(a)));
            case "EXP": return Need1(a => Num(Math.Exp(a)));
            case "LN": return Need1(a => a <= 0 ? FValue.Error("#NUM!") : Num(Math.Log(a)));
            case "LOG10": return Need1(a => a <= 0 ? FValue.Error("#NUM!") : Num(Math.Log10(a)));
            case "LOG": return Need1(a => { double b = vals.Length > 1 ? N(1) ?? double.NaN : 10; return a <= 0 || b <= 0 || b == 1 ? FValue.Error("#NUM!") : Num(Math.Log(a) / Math.Log(b)); });
            case "SIGN": return Need1(a => FValue.Num(Math.Sign(a)));
            case "TRUNC": return Need1(a => { double s = Math.Pow(10, vals.Length > 1 ? Math.Truncate(N(1) ?? 0) : 0); return Num(Math.Truncate(a * s) / s); });
            case "CEILING":
            case "CEILING.MATH":
            case "FLOOR":
            case "FLOOR.MATH":
                return Need1(a =>
                {
                    bool math = name.EndsWith(".MATH", StringComparison.Ordinal), up = name.StartsWith("CEILING", StringComparison.Ordinal);
                    double sig = vals.Length > 1 ? N(1) ?? 1 : 1;
                    if (sig == 0) return FValue.Num(0);
                    if (!math && a > 0 && sig < 0) return FValue.Error("#NUM!");
                    // Negative numbers: .MATH with a non-zero mode, or the old functions with a negative significance, round on the magnitude.
                    bool onMagnitude = a < 0 && (math ? vals.Length > 2 && (N(2) ?? 0) != 0 : sig < 0);
                    double s = Math.Abs(sig), q = Math.Round((onMagnitude ? Math.Abs(a) : a) / s, 12);
                    double r = (up ? Math.Ceiling(q) : Math.Floor(q)) * s;
                    return Num(onMagnitude ? -r : r);
                });
            case "MROUND": return Need2((a, m) => m == 0 ? FValue.Num(0) : Math.Sign(a) * Math.Sign(m) < 0 ? FValue.Error("#NUM!") : Num(Math.Round(a / m, MidpointRounding.AwayFromZero) * m));
            case "EVEN": return Need1(a => { double c = Math.Ceiling(Math.Abs(a)); if (c % 2 != 0) c++; return Num(Math.Sign(a) * c); });
            case "ODD": return Need1(a => { double c = Math.Ceiling(Math.Abs(a)); if (c % 2 == 0) c++; return Num(a < 0 ? -c : c); });
            case "FACT": return Need1(a => { if (a < 0) return FValue.Error("#NUM!"); double f = 1; for (int i = 2; i <= (int)a; i++) f *= i; return Num(f); });
            case "GCD":
            case "LCM":
            {
                if (numbers.Any(x => x < 0)) return FValue.Error("#NUM!");
                long acc = name == "GCD" ? 0 : 1;
                foreach (var x in numbers.Select(v => (long)Math.Truncate(v))) acc = name == "GCD" ? Gcd(acc, x) : (x == 0 || acc == 0 ? 0 : acc / Gcd(acc, x) * x);
                return FValue.Num(acc);
            }
            case "RANDBETWEEN": return Need2((a, b) => { long lo = (long)Math.Ceiling(a), hi = (long)Math.Floor(b); return lo > hi ? FValue.Error("#NUM!") : FValue.Num(Rng.NextInt64(lo, hi + 1)); });
            case "QUOTIENT": return Need2((a, b) => b == 0 ? FValue.Error("#DIV/0!") : FValue.Num(Math.Truncate(a / b)));
            case "RADIANS": return Need1(a => Num(a * Math.PI / 180));
            case "DEGREES": return Need1(a => Num(a * 180 / Math.PI));
            case "SIN": return Need1(a => Num(Math.Sin(a)));
            case "COS": return Need1(a => Num(Math.Cos(a)));
            case "TAN": return Need1(a => Num(Math.Tan(a)));
            case "ASIN": return Need1(a => a < -1 || a > 1 ? FValue.Error("#NUM!") : Num(Math.Asin(a)));
            case "ACOS": return Need1(a => a < -1 || a > 1 ? FValue.Error("#NUM!") : Num(Math.Acos(a)));
            case "ATAN": return Need1(a => Num(Math.Atan(a)));
            case "ATAN2": return Need2((x, y) => x == 0 && y == 0 ? FValue.Error("#DIV/0!") : Num(Math.Atan2(y, x)));

            // statistics
            case "MEDIAN":
            {
                if (numbers.Length == 0) return FValue.Error("#NUM!");
                var s = numbers.OrderBy(x => x).ToArray(); int m = s.Length / 2;
                return FValue.Num(s.Length % 2 == 1 ? s[m] : (s[m - 1] + s[m]) / 2);
            }
            case "MODE":
            case "MODE.SNGL":
            {
                var best = numbers.GroupBy(x => x).Where(g => g.Count() > 1).OrderByDescending(g => g.Count()).ThenBy(g => Array.IndexOf(numbers, g.Key)).FirstOrDefault();
                return best == null ? FValue.Error("#N/A") : FValue.Num(best.Key);
            }
            case "STDEV":
            case "STDEV.S":
            case "VAR":
            case "VAR.S":
            case "STDEV.P":
            case "STDEVP":
            case "VAR.P":
            case "VARP":
            {
                bool population = name.EndsWith("P", StringComparison.Ordinal); int n = numbers.Length;
                if (n - (population ? 0 : 1) <= 0) return FValue.Error("#DIV/0!");
                double mean = numbers.Average(), ss = numbers.Sum(x => (x - mean) * (x - mean)), variance = ss / (n - (population ? 0 : 1));
                return Num(name.StartsWith("STDEV", StringComparison.Ordinal) ? Math.Sqrt(variance) : variance);
            }

            // text
            case "FIND":
            case "SEARCH":
            {
                string find = S(0), within = S(1); int start = vals.Length > 2 ? (int)(N(2) ?? 1) : 1;
                if (start < 1 || start > within.Length + 1) return FValue.Error("#VALUE!");
                int at;
                if (name == "FIND") at = within.IndexOf(find, start - 1, StringComparison.Ordinal);
                else
                {
                    string pattern = WildcardRegex(find);
                    var m = System.Text.RegularExpressions.Regex.Match(within[(start - 1)..], pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant);
                    at = m.Success ? m.Index + start - 1 : -1;
                }
                return at < 0 ? FValue.Error("#VALUE!") : FValue.Num(at + 1);
            }
            case "SUBSTITUTE":
            {
                string text = S(0), old = S(1), nu = S(2);
                if (old.Length == 0) return FValue.Str(text);
                if (vals.Length < 4) return FValue.Str(text.Replace(old, nu, StringComparison.Ordinal));
                int which = (int)(N(3) ?? 0); if (which < 1) return FValue.Error("#VALUE!");
                int pos = -1; for (int i = 0; i < which; i++) { pos = text.IndexOf(old, pos + 1, StringComparison.Ordinal); if (pos < 0) return FValue.Str(text); }
                return FValue.Str(text[..pos] + nu + text[(pos + old.Length)..]);
            }
            case "REPLACE":
            {
                string text = S(0); if (N(1) is not double st || N(2) is not double len || st < 1 || len < 0) return FValue.Error("#VALUE!");
                int s0 = Math.Min((int)st - 1, text.Length), l = Math.Min((int)len, text.Length - s0);
                return FValue.Str(text[..s0] + S(3) + text[(s0 + l)..]);
            }
            case "REPT": return N(1) is double times && times >= 0 && S(0).Length * times <= 32767 ? FValue.Str(string.Concat(Enumerable.Repeat(S(0), (int)times))) : FValue.Error("#VALUE!");
            case "EXACT": return FValue.Bool(string.Equals(S(0), S(1), StringComparison.Ordinal));
            case "PROPER":
            {
                var sb = new StringBuilder(S(0).Length); bool up = true;
                foreach (char ch in S(0)) { sb.Append(up ? char.ToUpperInvariant(ch) : char.ToLowerInvariant(ch)); up = !char.IsLetter(ch); }
                return FValue.Str(sb.ToString());
            }
            case "CHAR": return Need1(a => a < 1 || a > 255 ? FValue.Error("#VALUE!") : FValue.Str(Encoding.Latin1.GetString(new[] { (byte)a })));
            case "UNICHAR": return Need1(a => a < 1 || a > 0x10FFFF ? FValue.Error("#VALUE!") : FValue.Str(char.ConvertFromUtf32((int)a)));
            case "CODE":
            case "UNICODE": { string s = S(0); return s.Length == 0 ? FValue.Error("#VALUE!") : FValue.Num(name == "CODE" && s[0] < 256 ? s[0] : char.ConvertToUtf32(s, 0)); }
            case "CLEAN": return FValue.Str(new string(S(0).Where(ch => ch >= 32).ToArray()));
            case "T": return vals.Length > 0 && vals[0].Kind == ValueKind.Text ? vals[0] : FValue.Str("");
            case "N": return vals.Length > 0 && vals[0].Kind is ValueKind.Number or ValueKind.Bool ? FValue.Num(vals[0].Number) : FValue.Num(0);
            case "FIXED":
                return Need1(a =>
                {
                    int dec = vals.Length > 1 ? (int)(N(1) ?? 2) : 2; bool noCommas = vals.Length > 2 && (N(2) ?? 0) != 0;
                    double r = Math.Round(a, Math.Clamp(dec, 0, 15), MidpointRounding.AwayFromZero);
                    if (dec < 0) { double s = Math.Pow(10, -dec); r = Math.Round(a / s, MidpointRounding.AwayFromZero) * s; dec = 0; }
                    return FValue.Str(r.ToString((noCommas ? "F" : "N") + Math.Max(0, dec).ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture));
                });
            case "TEXTBEFORE":
            case "TEXTAFTER":
            {
                string text = S(0), delim = S(1); int inst = vals.Length > 2 ? (int)(N(2) ?? 1) : 1;
                if (delim.Length == 0 || inst == 0) return FValue.Error("#VALUE!");
                int pos = -1;
                if (inst > 0) { for (int i = 0; i < inst; i++) { pos = text.IndexOf(delim, pos + 1, StringComparison.Ordinal); if (pos < 0) break; } }
                else { pos = text.Length; for (int i = 0; i < -inst; i++) { pos = pos <= 0 ? -1 : text.LastIndexOf(delim, pos - 1, StringComparison.Ordinal); if (pos < 0) break; } }
                if (pos < 0) return FValue.Error("#N/A");
                return FValue.Str(name == "TEXTBEFORE" ? text[..pos] : text[(pos + delim.Length)..]);
            }
            case "NUMBERVALUE":
            {
                string s = S(0).Replace(" ", ""); string dec = vals.Length > 1 ? S(1) : CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, grp = vals.Length > 2 ? S(2) : CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
                if (grp.Length > 0) s = s.Replace(grp[..1], ""); if (dec.Length > 0) s = s.Replace(dec[..1], ".");
                bool pct = s.EndsWith('%'); if (pct) s = s.TrimEnd('%');
                return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? FValue.Num(pct ? v / 100 : v) : FValue.Error("#VALUE!");
            }

            case "ISEVEN": return Need1(a => FValue.Bool(Math.Truncate(a) % 2 == 0));
            case "ISODD": return Need1(a => FValue.Bool(Math.Abs(Math.Truncate(a) % 2) == 1));

            // logic
            case "XOR": return FValue.Bool(vals.Select(Number).Count(v => v is double d && d != 0) % 2 == 1);

            // dates and times
            case "HOUR":
            case "MINUTE":
            case "SECOND":
                return Need1(a =>
                {
                    if (a < 0) return FValue.Error("#NUM!");
                    long secs = (long)Math.Round((a - Math.Floor(a)) * 86400, MidpointRounding.AwayFromZero) % 86400;
                    return FValue.Num(name == "HOUR" ? secs / 3600 : name == "MINUTE" ? secs / 60 % 60 : secs % 60);
                });
            case "TIME":
                if (N(0) is double h && N(1) is double mi && N(2) is double se)
                {
                    double t = (Math.Truncate(h) * 3600 + Math.Truncate(mi) * 60 + Math.Truncate(se)) / 86400;
                    return t < 0 ? FValue.Error("#NUM!") : FValue.Num(t - Math.Floor(t));
                }
                return FValue.Error("#VALUE!");
            case "DAYS": return Need2((end, start) => FValue.Num(Math.Floor(end) - Math.Floor(start)));
            case "DATEVALUE":
            {
                string s = S(0);
                if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out var dt) || DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                    return FValue.Num(Math.Floor(dt.ToOADate()));
                return FValue.Error("#VALUE!");
            }
            case "DATEDIF":
            {
                if (N(0) is not double a || N(1) is not double b) return FValue.Error("#VALUE!");
                if (a > b) return FValue.Error("#NUM!");
                DateTime d1, d2; try { d1 = DateTime.FromOADate(Math.Floor(a)); d2 = DateTime.FromOADate(Math.Floor(b)); } catch (ArgumentException) { return FValue.Error("#NUM!"); }
                int months = (d2.Year - d1.Year) * 12 + d2.Month - d1.Month; if (d2.Day < d1.Day) months--;
                return S(2).ToUpperInvariant() switch
                {
                    "D" => FValue.Num((d2 - d1).TotalDays),
                    "M" => FValue.Num(months),
                    "Y" => FValue.Num(months / 12),
                    "YM" => FValue.Num(months % 12),
                    "MD" => FValue.Num(d2.Day >= d1.Day ? d2.Day - d1.Day : (d2 - d2.AddMonths(-1).AddDays(d1.Day - d2.AddMonths(-1).Day)).TotalDays),
                    "YD" => FValue.Num((d2 - d1.AddYears(d2.Year - d1.Year - (new DateTime(d2.Year, d1.Month, Math.Min(d1.Day, DateTime.DaysInMonth(d2.Year, d1.Month))) > d2 ? 1 : 0))).TotalDays),
                    _ => FValue.Error("#NUM!"),
                };
            }
            case "WORKDAY":
            {
                if (N(0) is not double start || N(1) is not double days) return FValue.Error("#VALUE!");
                var holidays = vals.Skip(2).Where(v => v.Kind == ValueKind.Number).Select(v => Math.Floor(v.Number)).ToHashSet();
                double d = Math.Floor(start); int left = (int)Math.Truncate(days), step = left >= 0 ? 1 : -1;
                while (left != 0)
                {
                    d += step; var dow = DateTime.FromOADate(d).DayOfWeek;
                    if (dow is DayOfWeek.Saturday or DayOfWeek.Sunday || holidays.Contains(d)) continue;
                    left -= step;
                }
                return FValue.Num(d);
            }
            case "YEARFRAC":
            {
                if (N(0) is not double a || N(1) is not double b) return FValue.Error("#VALUE!");
                if (a > b) (a, b) = (b, a);
                int basis = vals.Length > 2 ? (int)(N(2) ?? 0) : 0;
                var d1 = DateTime.FromOADate(Math.Floor(a)); var d2 = DateTime.FromOADate(Math.Floor(b));
                return basis switch
                {
                    1 => FValue.Num((b - a) / (DateTime.IsLeapYear(d1.Year) && d1.Year == d2.Year ? 366 : 365.25)),
                    2 => FValue.Num((b - a) / 360),
                    3 => FValue.Num((b - a) / 365),
                    0 or 4 => FValue.Num(Days360(d1, d2, basis == 4) / 360.0),
                    _ => FValue.Error("#NUM!"),
                };
            }

            // finance
            case "PMT":
            case "PV":
            case "FV":
            case "NPER":
            {
                double rate = N(0) ?? double.NaN, a1 = N(1) ?? double.NaN, a2 = N(2) ?? double.NaN, a3 = vals.Length > 3 ? N(3) ?? 0 : 0; int type = vals.Length > 4 && (N(4) ?? 0) != 0 ? 1 : 0;
                if (double.IsNaN(rate) || double.IsNaN(a1) || double.IsNaN(a2)) return FValue.Error("#VALUE!");
                switch (name)
                {
                    case "PMT": { double nper = a1, pv = a2, fv = a3; return Num(rate == 0 ? -(pv + fv) / nper : -(rate * (fv + pv * Math.Pow(1 + rate, nper))) / ((1 + rate * type) * (Math.Pow(1 + rate, nper) - 1))); }
                    case "PV": { double nper = a1, pmt = a2, fv = a3; return Num(rate == 0 ? -(fv + pmt * nper) : -(fv + pmt * (1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / rate) / Math.Pow(1 + rate, nper)); }
                    case "FV": { double nper = a1, pmt = a2, pv = a3; return Num(rate == 0 ? -(pv + pmt * nper) : -(pv * Math.Pow(1 + rate, nper) + pmt * (1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / rate)); }
                    default:
                    {
                        double pmt = a1, pv = a2, fv = a3;
                        if (rate == 0) return pmt == 0 ? FValue.Error("#NUM!") : Num(-(pv + fv) / pmt);
                        double z = pmt * (1 + rate * type) / rate; double num = -fv + z, den = pv + z;
                        return num / den <= 0 ? FValue.Error("#NUM!") : Num(Math.Log(num / den) / Math.Log(1 + rate));
                    }
                }
            }
            case "IPMT":
            case "PPMT":
            {
                if (N(0) is not double rate || N(1) is not double per || N(2) is not double nper || N(3) is not double pv) return FValue.Error("#VALUE!");
                double fv = vals.Length > 4 ? N(4) ?? 0 : 0; int type = vals.Length > 5 && (N(5) ?? 0) != 0 ? 1 : 0;
                if (per < 1 || per > nper) return FValue.Error("#NUM!");
                double pmt = rate == 0 ? -(pv + fv) / nper : -(rate * (fv + pv * Math.Pow(1 + rate, nper))) / ((1 + rate * type) * (Math.Pow(1 + rate, nper) - 1));
                double ipmt;
                if (per == 1 && type == 1) ipmt = 0;
                else
                {
                    double k = type == 1 ? per - 2 : per - 1;
                    double balance = rate == 0 ? pv + pmt * k : pv * Math.Pow(1 + rate, k) + pmt * (1 + rate * type) * (Math.Pow(1 + rate, k) - 1) / rate;
                    ipmt = -balance * rate;
                }
                return Num(name == "IPMT" ? ipmt : pmt - ipmt);
            }
            case "NPV":
            {
                if (N(0) is not double rate) return FValue.Error("#VALUE!");
                double npv = 0; int t = 1;
                foreach (var v in vals.Skip(1)) { if (v.Kind != ValueKind.Number) continue; npv += v.Number / Math.Pow(1 + rate, t++); }
                return Num(npv);
            }
        }
        return FValue.Unsupported;
    }

    static long Gcd(long a, long b) { while (b != 0) (a, b) = (b, a % b); return Math.Abs(a); }

    static int Days360(DateTime a, DateTime b, bool european)
    {
        int d1 = a.Day, d2 = b.Day;
        if (european) { if (d1 == 31) d1 = 30; if (d2 == 31) d2 = 30; }
        else
        {
            bool lastFeb(DateTime d) => d.Month == 2 && d.Day == DateTime.DaysInMonth(d.Year, 2);
            if (lastFeb(a)) { if (lastFeb(b)) d2 = 30; d1 = 30; }
            if (d2 == 31 && d1 >= 30) d2 = 30;
            if (d1 == 31) d1 = 30;
        }
        return (b.Year - a.Year) * 360 + (b.Month - a.Month) * 30 + d2 - d1;
    }

    static bool Same(FValue a, FValue b)
    {
        if (a.Kind == ValueKind.Text || b.Kind == ValueKind.Text) return a.Kind == b.Kind && string.Equals(a.Text, b.Text, StringComparison.OrdinalIgnoreCase);
        return Number(a) == Number(b);
    }

    static string WildcardRegex(string pattern)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < pattern.Length; i++)
        {
            char ch = pattern[i];
            if (ch == '~' && i + 1 < pattern.Length) sb.Append(System.Text.RegularExpressions.Regex.Escape(pattern[++i].ToString()));
            else sb.Append(ch == '*' ? ".*" : ch == '?' ? "." : System.Text.RegularExpressions.Regex.Escape(ch.ToString()));
        }
        return sb.ToString();
    }
}
