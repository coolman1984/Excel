// Shows a value the way an Excel number format would.
// Anything it cannot do is reported as unsupported (null), never guessed.
using System.Drawing;
using System.Globalization;
using System.Text;

namespace ExcelViewer.Core;

public readonly record struct Shown(string Text, Color? Color);

/// <summary>Month and day names as Excel itself writes them. Windows can differ, such as Sept and Sep.</summary>
public sealed class DateNames { public string[] MonthAbbr, MonthFull, DayAbbr, DayFull; }

public sealed class NumFmt
{
    enum K { Lit, D0, DH, DQ, Dot, Comma, Pct, Exp, Slash, At, General, Y, Mo, Da, H, Mi, S, AmPm, EH, EM, ES, SecFrac, DateSep, TimeSep }

    sealed class Tok { public K Kind; public string Text = ""; public int Len; }

    sealed class Sec
    {
        public List<Tok> Toks = new();
        public string CondOp; public double CondVal; public Color? Color;
        public CultureInfo Culture; public bool Unsupported;
        public bool IsDate, IsText, IsGeneral;
        public bool Thousands; public int Scale, PctCount;
    }

    readonly List<Sec> numeric = new();
    readonly Sec textSec;
    readonly bool date1904; readonly DateNames names;
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public string Code { get; }
    public bool HasTextHandling => textSec != null;
    public bool IsDateFormat => numeric.Any(s => s.IsDate);
    public bool UsesNames => numeric.Any(s => s.IsDate && s.Toks.Any(t => t.Kind is K.Mo or K.Da && t.Len >= 3));

    public NumFmt(string code, bool date1904 = false, DateNames names = null)
    {
        Code = code ?? "General"; this.date1904 = date1904; this.names = names;
        var parts = Split(Code);
        var all = parts.Select(Parse).ToList();
        for (int i = 0; i < all.Count; i++)
        {
            if (i == 3 || (all[i].IsText && all.Count == 1)) textSec = all[i];
            else if (i < 3) numeric.Add(all[i]);
        }
    }

    // ---------- Splitting and reading the format text ----------

    static List<string> Split(string code)
    {
        var parts = new List<string>(); var sb = new StringBuilder(); bool q = false, br = false;
        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            if (q) { sb.Append(c); if (c == '"') q = false; continue; }
            if (br) { sb.Append(c); if (c == ']') br = false; continue; }
            if (c == '\\' || c == '_' || c == '*') { sb.Append(c); if (i + 1 < code.Length) sb.Append(code[++i]); continue; }
            if (c == '"') { q = true; sb.Append(c); continue; }
            if (c == '[') { br = true; sb.Append(c); continue; }
            if (c == ';') { parts.Add(sb.ToString()); sb.Clear(); continue; }
            sb.Append(c);
        }
        parts.Add(sb.ToString());
        return parts;
    }

    static Sec Parse(string s)
    {
        var sec = new Sec(); var t = sec.Toks;
        void Lit(string x) { if (t.Count > 0 && t[^1].Kind == K.Lit) t[^1].Text += x; else t.Add(new Tok { Kind = K.Lit, Text = x }); }
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i], lo = char.ToLowerInvariant(c);
            switch (c)
            {
                case '"': { int e = s.IndexOf('"', i + 1); if (e < 0) e = s.Length; Lit(s.Substring(i + 1, e - i - 1)); i = e; continue; }
                case '\\': if (i + 1 < s.Length) Lit(s[++i].ToString()); continue;
                case '_': if (i + 1 < s.Length) { i++; Lit(" "); } continue;
                case '*': i++; continue;
                case '[':
                    {
                        int e = s.IndexOf(']', i + 1); if (e < 0) e = s.Length - 1;
                        Bracket(s.Substring(i + 1, e - i - 1), sec, Lit); i = e; continue;
                    }
                case '0': t.Add(new Tok { Kind = K.D0 }); continue;
                case '#': t.Add(new Tok { Kind = K.DH }); continue;
                case '?': t.Add(new Tok { Kind = K.DQ }); continue;
                case '.': t.Add(new Tok { Kind = K.Dot }); continue;
                case ',': t.Add(new Tok { Kind = K.Comma }); continue;
                case '%': t.Add(new Tok { Kind = K.Pct }); continue;
                case '/': t.Add(new Tok { Kind = K.Slash }); continue;
                case ':': t.Add(new Tok { Kind = K.TimeSep }); continue;
                case '@': t.Add(new Tok { Kind = K.At }); continue;
            }
            if ((lo == 'e') && i + 1 < s.Length && (s[i + 1] == '+' || s[i + 1] == '-')) { t.Add(new Tok { Kind = K.Exp, Text = s[i + 1].ToString() }); i++; continue; }
            if (string.Compare(s, i, "General", 0, 7, StringComparison.OrdinalIgnoreCase) == 0) { t.Add(new Tok { Kind = K.General }); i += 6; continue; }
            if (string.Compare(s, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0) { t.Add(new Tok { Kind = K.AmPm, Text = s.Substring(i, 5) }); i += 4; continue; }
            if (string.Compare(s, i, "A/P", 0, 3, StringComparison.OrdinalIgnoreCase) == 0) { t.Add(new Tok { Kind = K.AmPm, Text = s.Substring(i, 3) }); i += 2; continue; }
            if (lo is 'y' or 'm' or 'd' or 'h' or 's')
            {
                int n = 1; while (i + n < s.Length && char.ToLowerInvariant(s[i + n]) == lo) n++;
                K k = lo switch { 'y' => K.Y, 'm' => K.Mo, 'd' => K.Da, 'h' => K.H, _ => K.S };
                t.Add(new Tok { Kind = k, Len = n }); i += n - 1; continue;
            }
            if (lo is 'g' or 'b' or 'e') { sec.Unsupported = true; continue; } // eras and other calendars
            Lit(c.ToString());
        }
        Resolve(sec);
        return sec;
    }

    static void Bracket(string b, Sec sec, Action<string> lit)
    {
        string lo = b.ToLowerInvariant();
        if (lo.Length > 0 && lo.All(ch => ch == 'h')) { sec.Toks.Add(new Tok { Kind = K.EH, Len = lo.Length }); return; }
        if (lo.Length > 0 && lo.All(ch => ch == 'm')) { sec.Toks.Add(new Tok { Kind = K.EM, Len = lo.Length }); return; }
        if (lo.Length > 0 && lo.All(ch => ch == 's')) { sec.Toks.Add(new Tok { Kind = K.ES, Len = lo.Length }); return; }
        Color? col = lo switch { "black" => Color.Black, "blue" => Color.Blue, "cyan" => Color.DarkCyan, "green" => Color.Green, "magenta" => Color.Magenta, "red" => Color.Red, "white" => Color.White, "yellow" => Color.Goldenrod, _ => null };
        if (col != null) { sec.Color = col; return; }
        if (lo.StartsWith("color")) return;
        foreach (var op in new[] { ">=", "<=", "<>", ">", "<", "=" })
            if (lo.StartsWith(op) && double.TryParse(lo[op.Length..], NumberStyles.Float, Inv, out double v)) { sec.CondOp = op; sec.CondVal = v; return; }
        if (b.StartsWith('$'))
        {
            int dash = b.IndexOf('-');
            lit(dash < 0 ? b[1..] : b[1..dash]);
            if (dash >= 0 && int.TryParse(b[(dash + 1)..], NumberStyles.HexNumber, Inv, out int id))
            {
                if (id > 0xFFFF) { sec.Unsupported = true; return; } // calendar or digit-shape flags
                if (id != 0 && id < 0xF000) { try { var c = new CultureInfo(id, false); if (c.Calendar is GregorianCalendar) sec.Culture = c; else sec.Unsupported = true; } catch (CultureNotFoundException) { sec.Unsupported = true; } }
            }
            return;
        }
        sec.Unsupported = true; // [DBNum1] and anything else unknown
    }

    static void Resolve(Sec sec)
    {
        var t = sec.Toks;
        sec.IsDate = t.Any(x => x.Kind is K.Y or K.Mo or K.Da or K.H or K.S or K.EH or K.EM or K.ES or K.AmPm);
        sec.IsGeneral = t.Any(x => x.Kind == K.General);
        sec.IsText = t.Any(x => x.Kind == K.At);
        if (sec.IsDate)
        {
            for (int i = 0; i < t.Count; i++)
            {
                if (t[i].Kind == K.Slash) t[i].Kind = K.DateSep;
                if (t[i].Kind == K.Mo)
                {
                    int p = i - 1; while (p >= 0 && !IsDatePart(t[p].Kind)) p--;
                    int n = i + 1; while (n < t.Count && !IsDatePart(t[n].Kind)) n++;
                    bool afterHour = p >= 0 && t[p].Kind is K.H or K.EH;
                    bool beforeSec = n < t.Count && t[n].Kind is K.S or K.ES;
                    if ((afterHour || beforeSec) && t[i].Len <= 2) t[i].Kind = K.Mi;
                }
                if (t[i].Kind == K.Dot && i + 1 < t.Count && t[i + 1].Kind == K.D0)
                {
                    int n = 0; while (i + 1 + n < t.Count && t[i + 1 + n].Kind == K.D0) n++;
                    t[i].Kind = K.SecFrac; t[i].Len = n; t.RemoveRange(i + 1, n);
                }
                else if (t[i].Kind is K.Dot or K.Comma or K.Pct) { t[i].Text = t[i].Kind == K.Dot ? "." : t[i].Kind == K.Comma ? "," : "%"; t[i].Kind = K.Lit; } // plain characters in a date
                else if (t[i].Kind is K.D0 or K.DH or K.DQ) sec.Unsupported = true;
            }
            return;
        }
        int last = -1, firstDigit = -1, dot = -1;
        for (int i = 0; i < t.Count; i++)
        {
            if (IsDigit(t[i].Kind)) { last = i; if (firstDigit < 0 && dot < 0) firstDigit = i; }
            if (t[i].Kind == K.Dot && dot < 0) dot = i;
            if (t[i].Kind == K.Pct) sec.PctCount++;
        }
        if (dot >= 0 && firstDigit < 0) firstDigit = dot;
        for (int i = 0; i < t.Count; i++)
        {
            if (t[i].Kind != K.Comma) continue;
            bool digitBefore = t.Take(i).Any(x => IsDigit(x.Kind)), digitAfterBeforeDot = t.Skip(i + 1).TakeWhile(x => x.Kind != K.Dot).Any(x => IsDigit(x.Kind));
            if (digitBefore && digitAfterBeforeDot) sec.Thousands = true;
            else if (digitBefore && i > last) sec.Scale++; // commas after the last digit divide by 1,000
            else if (!digitBefore && digitAfterBeforeDot) { }
            else { t[i].Kind = K.Lit; t[i].Text = ","; }
        }
        // A slash with no digits on both sides is a plain character.
        for (int i = 0; i < t.Count; i++)
            if (t[i].Kind == K.Slash && !(t.Take(i).Any(x => IsDigit(x.Kind)) && t.Skip(i + 1).Any(x => IsDigit(x.Kind) || (x.Kind == K.Lit && x.Text.Any(char.IsDigit)))))
            { t[i].Kind = K.Lit; t[i].Text = "/"; }
    }

    static bool IsDatePart(K k) => k is K.Y or K.Mo or K.Da or K.H or K.Mi or K.S or K.EH or K.EM or K.ES;
    static bool IsDigit(K k) => k is K.D0 or K.DH or K.DQ;

    // ---------- Showing a value ----------

    /// <summary>The shown text of a number. Null means this format cannot be carried.</summary>
    public Shown? Number(double v)
    {
        if (numeric.Count == 0) return new Shown(General(v), null);
        var (sec, abs, minus) = Choose(v);
        if (sec == null) return new Shown("#####", null);
        if (sec.Unsupported) return null;
        string body;
        if (sec.IsDate)
        {
            if (v < 0) return new Shown("#####", null);
            body = DateText(sec, v);
            if (body == null) return new Shown("#####", null);
        }
        else if (sec.IsGeneral) body = Render(sec, GeneralAbs(abs));
        else if (sec.Toks.Any(x => x.Kind == K.Slash)) body = Fraction(sec, abs);
        else if (sec.Toks.Any(x => x.Kind == K.Exp)) body = Scientific(sec, abs);
        else if (sec.Toks.Any(x => IsDigit(x.Kind) || x.Kind == K.Dot)) { body = Plain(sec, abs); if (minus && !body.Any(c => c is >= '1' and <= '9')) minus = false; } // Excel drops the minus when it rounds to zero
        else body = Render(sec, null);
        return new Shown((minus ? "-" : "") + body, sec.Color);
    }

    public string Text(string v)
    {
        if (textSec == null) return v;
        if (textSec.Unsupported) return null;
        var sb = new StringBuilder();
        foreach (var t in textSec.Toks) { if (t.Kind == K.At) sb.Append(v); else if (t.Kind == K.Lit) sb.Append(t.Text); else if (t.Kind == K.General) sb.Append(v); }
        return sb.ToString();
    }

    (Sec, double, bool) Choose(double v)
    {
        double abs = Math.Abs(v); int n = numeric.Count;
        if (numeric.Any(s => s.CondOp != null))
        {
            for (int i = 0; i < n; i++)
            {
                var s = numeric[i];
                if (s.CondOp != null ? Test(s, v) : i == n - 1 || (i == 1 && n == 3 && v < 0))
                    return (s, abs, v < 0 && !(s.CondOp is "<" or "<=" && s.CondVal <= 0));
            }
            return (null, abs, false);
        }
        if (n == 1) return (numeric[0], abs, v < 0);
        if (n == 2) return v >= 0 ? (numeric[0], abs, false) : (numeric[1], abs, false);
        return v > 0 ? (numeric[0], abs, false) : v < 0 ? (numeric[1], abs, false) : (numeric[2], abs, false);
    }

    static bool Test(Sec s, double v) => s.CondOp switch { ">" => v > s.CondVal, ">=" => v >= s.CondVal, "<" => v < s.CondVal, "<=" => v <= s.CondVal, "=" => v == s.CondVal, _ => v != s.CondVal };

    static string Render(Sec sec, string general)
    {
        var sb = new StringBuilder();
        foreach (var t in sec.Toks) { if (t.Kind == K.Lit) sb.Append(t.Text); else if (t.Kind == K.General) sb.Append(general); else if (t.Kind == K.Pct) sb.Append('%'); }
        return sb.ToString();
    }

    public static string General(double v) => (v < 0 ? "-" : "") + GeneralAbs(Math.Abs(v));

    // Excel's General shows at most 11 characters in a wide column.
    static string GeneralAbs(double a)
    {
        var nf = CultureInfo.CurrentCulture.NumberFormat;
        if (a == 0) return "0";
        if (a >= 1e11 || a < 1e-9) return a.ToString("0.#####E+00", CultureInfo.CurrentCulture);
        decimal d = (decimal)a;
        int intLen = a < 1 ? 1 : ((long)Math.Floor(a)).ToString(Inv).Length;
        d = Math.Round(d, Math.Clamp(10 - intLen, 0, 28), MidpointRounding.AwayFromZero);
        string s = d.ToString("0.############################", Inv);
        return s.Replace(".", nf.NumberDecimalSeparator);
    }

    static (decimal? d, double x) Scaled(Sec sec, double abs)
    {
        double x = abs;
        try
        {
            decimal d = (decimal)abs; // keeps 15 significant digits, as Excel does
            for (int i = 0; i < sec.PctCount; i++) d *= 100;
            for (int i = 0; i < sec.Scale; i++) d /= 1000;
            return (d, x);
        }
        catch (OverflowException)
        {
            for (int i = 0; i < sec.PctCount; i++) x *= 100;
            for (int i = 0; i < sec.Scale; i++) x /= 1000;
            return (null, x);
        }
    }

    static string Plain(Sec sec, double abs)
    {
        var t = sec.Toks; int dot = t.FindIndex(x => x.Kind == K.Dot);
        int decCount = dot < 0 ? 0 : t.Skip(dot + 1).Count(x => IsDigit(x.Kind));
        var (d, x) = Scaled(sec, abs);
        string fixedText = d != null
            ? Math.Round(d.Value, Math.Min(decCount, 28), MidpointRounding.AwayFromZero).ToString("F" + decCount, Inv)
            : x.ToString("F" + decCount, Inv);
        int p = fixedText.IndexOf('.');
        string intStr = (p < 0 ? fixedText : fixedText[..p]).TrimStart('0'), decStr = p < 0 ? "" : fixedText[(p + 1)..];
        return Assemble(sec, intStr, decStr, 0, dot < 0 ? t.Count : dot, dot);
    }

    // Puts digits into the placeholders.
    static string Assemble(Sec sec, string intStr, string decStr, int from, int intEnd, int dot, int until = -1)
    {
        var t = sec.Toks; var nf = CultureInfo.CurrentCulture.NumberFormat; if (until < 0) until = t.Count;
        var outp = new string[t.Count];
        int digitIx = intStr.Length - 1, placed = 0, firstPh = -1;
        for (int i = from; i < intEnd; i++) if (IsDigit(t[i].Kind)) { firstPh = i; break; }
        for (int i = intEnd - 1; i >= from; i--)
        {
            if (!IsDigit(t[i].Kind)) continue;
            var sb = new StringBuilder();
            void Put(char ch) { if (sec.Thousands && placed > 0 && placed % 3 == 0) sb.Insert(0, nf.NumberGroupSeparator); sb.Insert(0, ch); placed++; }
            if (digitIx >= 0) { Put(intStr[digitIx--]); }
            else if (t[i].Kind == K.D0) Put('0');
            else if (t[i].Kind == K.DQ) sb.Insert(0, ' ');
            if (i == firstPh) while (digitIx >= 0) Put(intStr[digitIx--]);
            outp[i] = sb.ToString();
        }
        string lead = firstPh < 0 && intStr.Length > 0 ? Grouped(intStr, sec.Thousands, nf) : "";
        if (dot >= 0)
        {
            outp[dot] = lead + nf.NumberDecimalSeparator; int n = 0;
            for (int i = dot + 1; i < until; i++)
            {
                if (!IsDigit(t[i].Kind)) continue;
                char ch = n < decStr.Length ? decStr[n] : '0'; bool restZero = n >= decStr.Length || decStr.AsSpan(n).TrimEnd('0').Length == 0; n++;
                outp[i] = t[i].Kind == K.D0 ? ch.ToString() : !restZero ? ch.ToString() : t[i].Kind == K.DQ ? " " : "";
            }
        }
        var res = new StringBuilder(lead.Length > 0 && dot < 0 ? lead : "");
        for (int i = from; i < until; i++)
        {
            if (t[i].Kind == K.Lit) res.Append(t[i].Text);
            else if (t[i].Kind == K.Pct) res.Append('%');
            else if (outp[i] != null) res.Append(outp[i]);
        }
        return res.ToString();
    }

    static string Grouped(string s, bool on, NumberFormatInfo nf)
    {
        if (!on) return s;
        var sb = new StringBuilder(); int c = 0;
        for (int i = s.Length - 1; i >= 0; i--) { if (c > 0 && c % 3 == 0) sb.Insert(0, nf.NumberGroupSeparator); sb.Insert(0, s[i]); c++; }
        return sb.ToString();
    }

    static string Scientific(Sec sec, double abs)
    {
        var t = sec.Toks; int ex = t.FindIndex(x => x.Kind == K.Exp), dot = t.FindIndex(x => x.Kind == K.Dot);
        if (dot > ex) dot = -1;
        int intPh = t.Take(dot < 0 ? ex : dot).Count(x => IsDigit(x.Kind));
        int decCount = dot < 0 ? 0 : t.Skip(dot + 1).Take(ex - dot - 1).Count(x => IsDigit(x.Kind));
        int expPh = t.Skip(ex + 1).Count(x => IsDigit(x.Kind));
        int e = abs == 0 ? 0 : (int)Math.Floor(Math.Log10(abs));
        int step = Math.Max(1, intPh);
        e = e - (((e % step) + step) % step);
        double m = abs == 0 ? 0 : abs / Math.Pow(10, e);
        double r = Math.Round(m, decCount, MidpointRounding.AwayFromZero);
        if (r >= Math.Pow(10, step)) { e += step; r = Math.Round(abs / Math.Pow(10, e), decCount, MidpointRounding.AwayFromZero); }
        string f = r.ToString("F" + decCount, Inv); int p = f.IndexOf('.');
        string head = Assemble(sec, (p < 0 ? f : f[..p]).TrimStart('0'), p < 0 ? "" : f[(p + 1)..], 0, dot < 0 ? ex : dot, dot, ex);
        string sign = e < 0 ? "-" : t[ex].Text == "+" ? "+" : "";
        return head + "E" + sign + Math.Abs(e).ToString(new string('0', Math.Max(1, expPh)), Inv) + string.Concat(t.Skip(ex + 1).Where(x => x.Kind == K.Lit).Select(x => x.Text));
    }

    static string Fraction(Sec sec, double abs)
    {
        var t = sec.Toks; int sl = t.FindIndex(x => x.Kind == K.Slash);
        int numEnd = sl; while (numEnd > 0 && !IsDigit(t[numEnd - 1].Kind)) numEnd--;
        int numStart = numEnd; while (numStart > 0 && IsDigit(t[numStart - 1].Kind)) numStart--;
        bool hasWhole = t.Take(numStart).Any(x => IsDigit(x.Kind));
        int denLen = t.Skip(sl + 1).Count(x => IsDigit(x.Kind)); string fixedDen = string.Concat(t.Skip(sl + 1).Where(x => x.Kind == K.Lit).SelectMany(x => x.Text).Where(char.IsDigit));
        double whole = hasWhole ? Math.Floor(abs) : 0, frac = abs - whole; long num, den;
        if (fixedDen.Length > 0) { den = long.Parse(fixedDen, Inv); num = (long)Math.Round(frac * den, MidpointRounding.AwayFromZero); }
        else (num, den) = BestFraction(frac, (long)Math.Pow(10, Math.Max(1, denLen)) - 1);
        if (hasWhole && num >= den && den > 0) { whole += 1; num = 0; }
        var sb = new StringBuilder();
        if (hasWhole) sb.Append(whole == 0 && num != 0 ? "" : whole.ToString("F0", Inv)).Append(' ');
        if (num != 0 || !hasWhole) sb.Append(num.ToString(Inv)).Append('/').Append(den.ToString(Inv));
        return sb.ToString().Trim();
    }

    static (long, long) BestFraction(double f, long maxDen)
    {
        if (f <= 0) return (0, 1);
        long pa = 0, qa = 1, pb = 1, qb = 0; double x = f;
        while (true)
        {
            double fl = Math.Floor(x); long a = (long)fl, p = a * pb + pa, q = a * qb + qa;
            if (q > maxDen) break;
            pa = pb; qa = qb; pb = p; qb = q;
            if (x - fl < 1e-12) break;
            x = 1 / (x - fl);
        }
        return qb == 0 ? (0, 1) : (pb, qb);
    }

    string DateText(Sec sec, double v)
    {
        if (v >= (date1904 ? 2957004 : 2958466)) return null;
        var t = sec.Toks; int fracDigits = t.FirstOrDefault(x => x.Kind == K.SecFrac)?.Len ?? 0;
        double p10 = Math.Pow(10, fracDigits);
        double totalSec = Math.Round(v * 86400 * p10, MidpointRounding.AwayFromZero) / p10;
        long days = (long)Math.Floor(totalSec / 86400); double secOfDay = totalSec - days * 86400.0;
        int y, mo, d, dow;
        if (date1904) { var dt = new DateTime(1904, 1, 1).AddDays(days); y = dt.Year; mo = dt.Month; d = dt.Day; dow = (int)dt.DayOfWeek; }
        else
        {
            dow = (int)((days + 6) % 7); // serial 1 counts as Sunday, the way Excel does
            if (days == 60) { y = 1900; mo = 2; d = 29; }
            else if (days == 0) { y = 1900; mo = 1; d = 0; }
            else { var dt = new DateTime(1899, 12, days < 60 ? 31 : 30).AddDays(days); y = dt.Year; mo = dt.Month; d = dt.Day; }
        }
        long whole = (long)Math.Floor(secOfDay); double fr = secOfDay - whole;
        int h = (int)(whole / 3600), mi = (int)(whole % 3600 / 60), s = (int)(whole % 60);
        bool ampm = t.Any(x => x.Kind == K.AmPm);
        var ci = sec.Culture ?? CultureInfo.CurrentCulture; var dtf = ci.DateTimeFormat;
        var nm = sec.Culture == null ? names : null; // a format naming its own language keeps that language
        var sb = new StringBuilder();
        foreach (var k in t)
            switch (k.Kind)
            {
                case K.Lit: sb.Append(k.Text); break;
                case K.DateSep: sb.Append(CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator); break;
                case K.TimeSep: sb.Append(CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator); break;
                case K.Y: sb.Append(k.Len <= 2 ? (y % 100).ToString("00", Inv) : y.ToString("0000", Inv)); break;
                case K.Mo:
                    sb.Append(k.Len switch { 1 => mo.ToString(Inv), 2 => mo.ToString("00", Inv), 3 => nm?.MonthAbbr[mo - 1] ?? dtf.AbbreviatedMonthNames[mo - 1], 5 => (nm?.MonthFull[mo - 1] ?? dtf.MonthNames[mo - 1])[..1], _ => nm?.MonthFull[mo - 1] ?? dtf.MonthNames[mo - 1] });
                    break;
                case K.Da:
                    sb.Append(k.Len switch { 1 => d.ToString(Inv), 2 => d.ToString("00", Inv), 3 => nm?.DayAbbr[dow] ?? dtf.AbbreviatedDayNames[dow], _ => nm?.DayFull[dow] ?? dtf.DayNames[dow] });
                    break;
                case K.H: { int hh = ampm ? (h % 12 == 0 ? 12 : h % 12) : h; sb.Append(k.Len == 1 ? hh.ToString(Inv) : hh.ToString("00", Inv)); break; }
                case K.Mi: sb.Append(k.Len == 1 ? mi.ToString(Inv) : mi.ToString("00", Inv)); break;
                case K.S: sb.Append(k.Len == 1 ? s.ToString(Inv) : s.ToString("00", Inv)); break;
                case K.SecFrac: sb.Append(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator).Append(Math.Round(fr * p10, MidpointRounding.AwayFromZero).ToString(new string('0', k.Len), Inv)); break;
                case K.AmPm:
                    if (k.Text.Length == 5) sb.Append(char.IsLower(k.Text[0]) ? (h < 12 ? "am" : "pm") : (h < 12 ? "AM" : "PM")); // as this format writes it
                    else sb.Append(h < 12 ? k.Text[0] : k.Text[2]);
                    break;
                case K.EH: sb.Append(((long)Math.Floor(totalSec / 3600)).ToString(new string('0', k.Len), Inv)); break;
                case K.EM: sb.Append(((long)Math.Floor(totalSec / 60)).ToString(new string('0', k.Len), Inv)); break;
                case K.ES: sb.Append(((long)Math.Floor(totalSec)).ToString(new string('0', k.Len), Inv)); break;
            }
        return sb.ToString();
    }
}
