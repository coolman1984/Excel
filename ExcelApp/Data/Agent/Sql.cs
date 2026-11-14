// A small read-only SQL for the store: SELECT ... FROM ... [JOIN ... ON ...] [WHERE] [GROUP BY] [HAVING] [ORDER BY] [LIMIT].
// It is translated into a QuerySpec and run by the same engine as the Data mode, so a question reads only the columns it uses.
using System.Globalization;
using System.Text;
using System.Text.Json;
using ExcelApp.Data.Query;

namespace ExcelApp.Data.Agent;

/// <summary>A parsed query ready for the engine, with the names to show for each output column.</summary>
public sealed record SqlPlan(QuerySpec Spec, IReadOnlyList<(string Internal, string Shown)>? Output, int? Limit);

public static class Sql
{
    // ---------------- tokens ----------------
    enum K { Ident, Quoted, Number, Text, Sym, End }
    readonly record struct Tok(K Kind, string Text, int At);

    static List<Tok> Lex(string s)
    {
        var t = new List<Tok>(); int i = 0;
        while (i < s.Length)
        {
            char c = s[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }
            if (c == '-' && i + 1 < s.Length && s[i + 1] == '-') { while (i < s.Length && s[i] != '\n') i++; continue; }
            int at = i;
            if (c == '\'')
            {
                var b = new StringBuilder(); i++;
                while (true)
                {
                    if (i >= s.Length) throw new UserError("A text value starting at character " + (at + 1) + " has no closing quote (').");
                    if (s[i] == '\'') { if (i + 1 < s.Length && s[i + 1] == '\'') { b.Append('\''); i += 2; continue; } i++; break; }
                    b.Append(s[i++]);
                }
                t.Add(new(K.Text, b.ToString(), at)); continue;
            }
            if (c is '"' or '[' or '`')
            {
                char close = c == '[' ? ']' : c; int end = s.IndexOf(close, i + 1);
                if (end < 0) throw new UserError("A name starting at character " + (at + 1) + " has no closing " + close + ".");
                t.Add(new(K.Quoted, s[(i + 1)..end], at)); i = end + 1; continue;
            }
            if (char.IsDigit(c) || (c == '.' && i + 1 < s.Length && char.IsDigit(s[i + 1])))
            {
                while (i < s.Length && (char.IsDigit(s[i]) || s[i] == '.')) i++;
                if (i < s.Length && s[i] is 'e' or 'E' && i + 1 < s.Length && (char.IsDigit(s[i + 1]) || s[i + 1] is '+' or '-')) { i += 2; while (i < s.Length && char.IsDigit(s[i])) i++; }
                t.Add(new(K.Number, s[at..i], at)); continue;
            }
            if (char.IsLetter(c) || c == '_')
            {
                while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '_')) i++;
                t.Add(new(K.Ident, s[at..i], at)); continue;
            }
            string two = i + 1 < s.Length ? s.Substring(i, 2) : "";
            if (two is "<>" or "!=" or "<=" or ">=") { t.Add(new(K.Sym, two == "!=" ? "<>" : two, at)); i += 2; continue; }
            if ("(),*+-/=<>.;%".Contains(c)) { t.Add(new(K.Sym, c.ToString(), at)); i++; continue; }
            throw new UserError($"The character '{c}' at position {at + 1} is not understood.");
        }
        t.Add(new(K.End, "", s.Length)); return t;
    }

    // ---------------- syntax tree ----------------
    abstract record E;
    sealed record Name(string? Table, string Col) : E;
    sealed record Num(double V) : E;
    sealed record Str(string V) : E;
    sealed record Bool(bool V) : E;
    sealed record Null : E;
    sealed record Bin(string Op, E L, E R) : E;
    sealed record Neg(E V) : E;
    sealed record Fn(string F, List<E> Args) : E;
    sealed record Agg(string F, bool Distinct, E? Arg) : E;   // Arg null = COUNT(*)
    sealed record Star : E;

    abstract record C;
    sealed record And(C L, C R) : C;
    sealed record Or(C L, C R) : C;
    sealed record Not(C V) : C;
    sealed record Cmp(string Op, E L, E R) : C;
    sealed record Like(E L, string Pattern) : C;
    sealed record In(E L, List<E> Values) : C;
    sealed record IsNull(E L) : C;
    sealed record Between(E L, E A, E B) : C;

    sealed record Item(E Expr, string? Alias, string Text);
    sealed record Source(string Name, string? Alias, string Kind = "inner", List<(E L, E R)>? On = null);
    sealed record Order(E Expr, bool Desc);
    sealed class Query
    {
        public bool Distinct; public List<Item> Items = new(); public Source From = null!; public List<Source> Joins = new();
        public C? Where; public List<E> Group = new(); public C? Having; public List<Order> Order = new(); public int? Limit;
    }

    static readonly HashSet<string> Aggs = new(StringComparer.OrdinalIgnoreCase) { "SUM", "COUNT", "AVG", "MIN", "MAX" };
    static readonly HashSet<string> Scalars = new(StringComparer.OrdinalIgnoreCase) { "ABS", "ROUND", "COALESCE", "MONTH", "PCT" };
    static readonly HashSet<string> Reserved = new(StringComparer.OrdinalIgnoreCase) { "SELECT", "DISTINCT", "FROM", "JOIN", "INNER", "LEFT", "ON", "WHERE", "GROUP", "BY", "HAVING", "ORDER", "LIMIT", "AND", "OR", "NOT", "AS", "ASC", "DESC", "IN", "IS", "NULL", "LIKE", "BETWEEN", "TRUE", "FALSE" };

    sealed class P
    {
        readonly List<Tok> t; readonly string src; int p;
        public P(string s) { src = s; t = Lex(s); }
        Tok Cur => t[p];
        bool Word(string w) => Cur.Kind == K.Ident && Cur.Text.Equals(w, StringComparison.OrdinalIgnoreCase);
        bool TakeWord(string w) { if (!Word(w)) return false; p++; return true; }
        bool Sym(string s) => Cur.Kind == K.Sym && Cur.Text == s;
        bool TakeSym(string s) { if (!Sym(s)) return false; p++; return true; }
        UserError Bad(string expected) => new(Cur.Kind == K.End ? $"The query ends too early: {expected} was expected." : $"Expected {expected} at \"{Near()}\" (character {Cur.At + 1}).");
        string Near() { var s = src[Cur.At..]; return s.Length > 25 ? s[..25] + "..." : s; }
        void NeedWord(string w) { if (!TakeWord(w)) throw Bad(w.ToUpperInvariant()); }
        void NeedSym(string s) { if (!TakeSym(s)) throw Bad("\"" + s + "\""); }

        public Query Parse()
        {
            if (!Word("SELECT")) throw new UserError("Only SELECT queries are allowed: the agent access reads the data and never changes it.");
            p++; var q = new Query { Distinct = TakeWord("DISTINCT") };
            do { int start = Cur.At; var e = Expr(); int end = Cur.At; string? alias = Alias(); q.Items.Add(new(e, alias, src[start..end].Trim())); } while (TakeSym(","));
            NeedWord("FROM"); q.From = Src();
            while (true)
            {
                string kind;
                if (TakeWord("JOIN")) kind = "inner";
                else if (Word("INNER")) { p++; NeedWord("JOIN"); kind = "inner"; }
                else if (Word("LEFT")) { p++; TakeWord("OUTER"); NeedWord("JOIN"); kind = "left"; }
                else break;
                var s = Src(); NeedWord("ON"); var on = new List<(E, E)>();
                do { var l = Expr(); NeedSym("="); on.Add((l, Expr())); } while (TakeWord("AND"));
                q.Joins.Add(s with { Kind = kind, On = on });
            }
            if (TakeWord("WHERE")) q.Where = Cond();
            if (TakeWord("GROUP")) { NeedWord("BY"); do q.Group.Add(Expr()); while (TakeSym(",")); }
            if (TakeWord("HAVING")) q.Having = Cond();
            if (TakeWord("ORDER")) { NeedWord("BY"); do { var e = Expr(); bool desc = TakeWord("DESC"); if (!desc) TakeWord("ASC"); q.Order.Add(new(e, desc)); } while (TakeSym(",")); }
            if (TakeWord("LIMIT")) { if (Cur.Kind != K.Number || !int.TryParse(Cur.Text, NumberStyles.None, CultureInfo.InvariantCulture, out int n)) throw Bad("a whole number after LIMIT"); p++; q.Limit = n; }
            TakeSym(";");
            if (Cur.Kind != K.End) throw Bad("the end of the query");
            return q;
        }
        string? Alias()
        {
            if (TakeWord("AS")) return NameText() ?? throw Bad("a name after AS");
            if ((Cur.Kind == K.Ident && !Reserved.Contains(Cur.Text)) || Cur.Kind == K.Quoted) return NameText();
            return null;
        }
        string? NameText() { if (Cur.Kind is K.Ident or K.Quoted) return t[p++].Text; return null; }
        Source Src()
        {
            var n = NameText() ?? throw Bad("a table name");
            return new(n, Alias());
        }

        // conditions: OR < AND < NOT < predicate
        C Cond() { var l = AndC(); while (TakeWord("OR")) l = new Or(l, AndC()); return l; }
        C AndC() { var l = NotC(); while (TakeWord("AND")) l = new And(l, NotC()); return l; }
        C NotC()
        {
            if (TakeWord("NOT")) return new Not(NotC());
            if (Sym("("))
            {
                int save = p; p++;
                try { var c = Cond(); if (TakeSym(")")) return c; } catch (UserError) { }
                p = save;   // "(" opened an expression such as (Sales - Cost) > 5
            }
            return Pred();
        }
        C Pred()
        {
            var l = Expr();
            if (TakeWord("IS")) { bool not = TakeWord("NOT"); NeedWord("NULL"); C c = new IsNull(l); return not ? new Not(c) : c; }
            bool neg = TakeWord("NOT");
            if (TakeWord("LIKE")) { if (Cur.Kind != K.Text) throw Bad("a quoted pattern after LIKE"); C c = new Like(l, t[p++].Text); return neg ? new Not(c) : c; }
            if (TakeWord("IN")) { NeedSym("("); var v = new List<E>(); do v.Add(Expr()); while (TakeSym(",")); NeedSym(")"); C c = new In(l, v); return neg ? new Not(c) : c; }
            if (TakeWord("BETWEEN")) { var a = Expr(); NeedWord("AND"); C c = new Between(l, a, Expr()); return neg ? new Not(c) : c; }
            if (neg) throw Bad("LIKE, IN or BETWEEN after NOT");
            foreach (var op in new[] { "=", "<>", "<=", ">=", "<", ">" }) if (TakeSym(op)) return new Cmp(op, l, Expr());
            throw Bad("a comparison (=, <>, <, >, <=, >=, LIKE, IN, BETWEEN or IS NULL)");
        }

        // expressions: + - < * / < unary < primary
        E Expr() { var l = Term(); while (true) { if (TakeSym("+")) l = new Bin("+", l, Term()); else if (TakeSym("-")) l = new Bin("-", l, Term()); else return l; } }
        E Term() { var l = Unary(); while (true) { if (TakeSym("*")) l = new Bin("*", l, Unary()); else if (TakeSym("/")) l = new Bin("/", l, Unary()); else return l; } }
        E Unary() { if (TakeSym("-")) return new Neg(Unary()); if (TakeSym("+")) return Unary(); return Primary(); }
        E Primary()
        {
            if (TakeSym("*")) return new Star();
            if (TakeSym("(")) { var e = Expr(); NeedSym(")"); return e; }
            if (Cur.Kind == K.Number) return new Num(double.Parse(t[p++].Text, NumberStyles.Float, CultureInfo.InvariantCulture));
            if (Cur.Kind == K.Text) return new Str(t[p++].Text);
            if (Cur.Kind == K.Ident && Cur.Text.Equals("NULL", StringComparison.OrdinalIgnoreCase)) { p++; return new Null(); }
            if (Cur.Kind == K.Ident && Cur.Text.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) { p++; return new Bool(true); }
            if (Cur.Kind == K.Ident && Cur.Text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) { p++; return new Bool(false); }
            if (Cur.Kind == K.Ident && Cur.Text.Equals("DATE", StringComparison.OrdinalIgnoreCase) && t[p + 1].Kind == K.Text) { p++; return new Str(t[p++].Text); }
            if (Cur.Kind == K.Ident && Reserved.Contains(Cur.Text)) throw Bad("a column, number or text");
            if (Cur.Kind is K.Ident or K.Quoted)
            {
                var first = t[p++];
                if (first.Kind == K.Ident && Sym("("))
                {
                    p++; string f = first.Text.ToUpperInvariant();
                    if (Aggs.Contains(f))
                    {
                        if (TakeSym("*")) { if (f != "COUNT") throw new UserError($"{f}(*) is not allowed; only COUNT(*)."); NeedSym(")"); return new Agg("COUNT", false, null); }
                        bool distinct = TakeWord("DISTINCT"); var arg = Expr(); NeedSym(")"); return new Agg(f, distinct, arg);
                    }
                    if (!Scalars.Contains(f)) throw new UserError($"The function {first.Text} is not available. Use SUM, COUNT, AVG, MIN, MAX, ABS, ROUND, COALESCE, MONTH or PCT.");
                    var args = new List<E>(); if (!TakeSym(")")) { do args.Add(Expr()); while (TakeSym(",")); NeedSym(")"); }
                    return new Fn(f, args);
                }
                if (TakeSym(".")) { var col = NameText() ?? throw Bad("a column name after the dot"); return new Name(first.Text, col); }
                return new Name(null, first.Text);
            }
            throw Bad("a column, number or text");
        }
    }

    // ---------------- translation to the engine ----------------
    /// <summary>Parses <paramref name="sql"/> against the store's tables. Every problem is a plain sentence (UserError).</summary>
    public static SqlPlan Plan(string sql, IReadOnlyList<TableInfo> tables)
    {
        if (string.IsNullOrWhiteSpace(sql)) throw new UserError("The query is empty.");
        var q = new P(sql).Parse();
        TableInfo Find(string name) =>
            tables.Where(x => x.Id.Equals(name, StringComparison.OrdinalIgnoreCase) || x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)).OrderBy(x => x.Kind == "report" ? 0 : 1).FirstOrDefault()
            ?? throw new UserError($"There is no table named \"{name}\". Tables: {string.Join(", ", tables.Select(x => x.Id).Distinct().OrderBy(x => x, StringComparer.OrdinalIgnoreCase))}.");
        var left = Find(q.From.Name);
        var leftNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { left.Id, left.Name }; if (q.From.Alias != null) leftNames.Add(q.From.Alias);
        static Dictionary<string, string> Canon(TableInfo t) { var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); foreach (var c in t.Columns) d.TryAdd(c.Name, c.Name); return d; }
        var leftCols = Canon(left);
        var joins = new List<(Source S, TableInfo Info, string Prefix, HashSet<string> Names, Dictionary<string, string> Cols)>();
        foreach (var j in q.Joins)
        {
            var info = Find(j.Name); var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { info.Id, info.Name }; if (j.Alias != null) names.Add(j.Alias);
            joins.Add((j, info, (j.Alias ?? info.Id) + ".", names, Canon(info)));
        }

        string AllColumns() => string.Join(", ", left.Columns.Where(c => !c.Name.StartsWith('_')).Select(c => c.Name).Concat(joins.SelectMany(j => j.Info.Columns.Where(c => !c.Name.StartsWith('_')).Select(c => j.Prefix + c.Name))));
        UserError Unknown(Name n) => new($"There is no column named \"{(n.Table != null ? n.Table + "." : "") + n.Col}\". Columns: {AllColumns()}.");
        string Col(Name n)
        {
            if (n.Table != null)
            {
                if (leftNames.Contains(n.Table)) return leftCols.TryGetValue(n.Col, out var lc) ? lc : throw Unknown(n);
                foreach (var j in joins) if (j.Names.Contains(n.Table)) return j.Cols.TryGetValue(n.Col, out var jc) ? j.Prefix + jc : throw Unknown(n);
                throw new UserError($"\"{n.Table}\" is not a table in this query.");
            }
            if (leftCols.TryGetValue(n.Col, out var c)) return c;
            foreach (var j in joins) if (j.Cols.TryGetValue(n.Col, out var jc)) return j.Prefix + jc;
            throw Unknown(n);
        }
        string RawCol(Name n, Dictionary<string, string> cols, string table) => cols.TryGetValue(n.Col, out var c) ? c : throw new UserError($"There is no column named \"{n.Col}\" in {table}.");

        var spec = new QuerySpec { Source = new(left.Kind, left.Id), Computed = new(), Aggregates = new(), GroupBy = new(), Sort = new() };
        foreach (var j in joins)
        {
            var keys = new List<JoinKey>();
            foreach (var (l, r) in j.S.On!)
            {
                if (l is not Name ln || r is not Name rn) throw new UserError("JOIN ... ON compares a column with a column, such as a.Region = b.Region.");
                // either side may name either table
                bool lRight = ln.Table != null && j.Names.Contains(ln.Table);
                var (lk, rk) = lRight ? (rn, ln) : (ln, rn);
                keys.Add(new JoinKey(Col(lk with { Table = lk.Table != null && leftNames.Contains(lk.Table) ? lk.Table : null }), RawCol(rk, j.Cols, j.Info.Id)));
            }
            spec.Joins ??= new(); spec.Joins.Add(new JoinSpec { Right = new(j.Info.Kind, j.Info.Id), On = keys, Kind = j.S.Kind, Prefix = j.Prefix });
        }

        int hidden = 0;
        string ExprText(E e) => e switch
        {
            Name n => "[" + Col(n) + "]",
            Num x => x.V.ToString("0.#################", CultureInfo.InvariantCulture),
            Bin b => "(" + ExprText(b.L) + " " + b.Op + " " + ExprText(b.R) + ")",
            Neg x => "-(" + ExprText(x.V) + ")",
            Fn f => f.F.ToLowerInvariant() + "(" + string.Join(", ", f.Args.Select(ExprText)) + ")",
            Str => throw new UserError("Text values cannot be used in a calculation."),
            Agg => throw new UserError("Calculations on top of SUM, COUNT, AVG, MIN or MAX are not supported yet; give the aggregate its own column."),
            _ => throw new UserError("That calculation is not supported.")
        };
        string Computed(E e, string? name = null) { name ??= "_e" + (++hidden); spec.Computed!.Add(new(name, ExprText(e))); return name; }
        string Show(E e) => e switch
        {
            Name n => (n.Table != null ? n.Table + "." : "") + n.Col,
            Agg a => a.F + "(" + (a.Distinct ? "DISTINCT " : "") + (a.Arg == null ? "*" : Show(a.Arg)) + ")",
            Num x => x.V.ToString(CultureInfo.InvariantCulture),
            Bin b => Show(b.L) + " " + b.Op + " " + Show(b.R),
            Neg x => "-" + Show(x.V),
            Fn f => f.F + "(" + string.Join(", ", f.Args.Select(Show)) + ")",
            Str s => "'" + s.V + "'",
            _ => "?"
        };
        bool HasAgg(E? e) => e switch { Agg => true, Bin b => HasAgg(b.L) || HasAgg(b.R), Neg n => HasAgg(n.V), Fn f => f.Args.Any(HasAgg), _ => false };
        bool CondAgg(C? c) => c switch { And a => CondAgg(a.L) || CondAgg(a.R), Or o => CondAgg(o.L) || CondAgg(o.R), Not n => CondAgg(n.V), Cmp m => HasAgg(m.L) || HasAgg(m.R), Between b => HasAgg(b.L), In i => HasAgg(i.L), IsNull n => HasAgg(n.L), Like l => HasAgg(l.L), _ => false };

        bool grouped = q.Distinct || q.Group.Count > 0 || q.Items.Any(i => HasAgg(i.Expr)) || CondAgg(q.Having);
        var aggNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);   // shown text -> engine column
        string AggCol(Agg a, string? alias)
        {
            string key = Show(a);
            if (alias == null && aggNames.TryGetValue(key, out var have)) return have;
            string name = alias ?? (spec.Aggregates!.Any(x => x.As.Equals(key, StringComparison.OrdinalIgnoreCase)) ? "_a" + (++hidden) : key);
            string? col = a.Arg switch { null => null, Name n => Col(n), Star => throw new UserError("Only COUNT(*) may use *."), var e => Computed(e) };
            string fn = a.F switch { "SUM" => "sum", "AVG" => "avg", "MIN" => "min", "MAX" => "max", _ => a.Distinct ? "countDistinct" : "count" };
            if (a.Distinct && a.F != "COUNT") throw new UserError("DISTINCT works inside COUNT only.");
            spec.Aggregates!.Add(new Aggregate(fn, col, name)); aggNames.TryAdd(key, name); return name;
        }

        var output = new List<(string Internal, string Shown)>(); bool star = false;
        // the name to show for a column: a joined table's column keeps its "table." prefix only when the left table has the same name
        string Plain(string c) { foreach (var j in joins) if (c.StartsWith(j.Prefix, StringComparison.OrdinalIgnoreCase)) { var raw = c[j.Prefix.Length..]; return leftCols.ContainsKey(raw) ? c : raw; } return c; }
        if (grouped)
        {
            var groupCols = (q.Distinct && q.Group.Count == 0 ? q.Items.Select(i => i.Expr) : q.Group).Select(e => e is Name n ? Col(n) : throw new UserError("GROUP BY and DISTINCT take column names.")).ToList();
            spec.GroupBy!.AddRange(groupCols);
            foreach (var item in q.Items)
            {
                if (item.Expr is Star) throw new UserError("SELECT * cannot be used with GROUP BY or aggregates.");
                if (item.Expr is Agg a) { output.Add((AggCol(a, item.Alias), item.Alias ?? Show(a))); continue; }
                if (item.Expr is Name n) { var c = Col(n); if (!groupCols.Contains(c, StringComparer.OrdinalIgnoreCase)) throw new UserError($"\"{Show(n)}\" must be in GROUP BY or inside SUM, COUNT, AVG, MIN or MAX."); output.Add((c, item.Alias ?? Plain(c))); continue; }
                throw new UserError("Calculations on top of SUM, COUNT, AVG, MIN or MAX are not supported yet; give the aggregate its own column.");
            }
        }
        else
        {
            foreach (var item in q.Items)
            {
                if (item.Expr is Star) { star = true; foreach (var c in left.Columns.Where(c => !c.Name.StartsWith('_'))) output.Add((c.Name, c.Name)); foreach (var j in joins) foreach (var c in j.Info.Columns.Where(c => !c.Name.StartsWith('_'))) output.Add((j.Prefix + c.Name, j.Prefix + c.Name)); continue; }
                if (item.Expr is Name n) { var c = Col(n); output.Add((c, item.Alias ?? Plain(c))); continue; }
                if (item.Expr is Str or Num or Bool or Null) throw new UserError("A fixed value in SELECT is not supported; select columns or calculations.");
                string shown = item.Alias ?? item.Text;
                output.Add((Computed(item.Expr, item.Alias ?? "_e" + (++hidden)), shown));
            }
        }
        // resolve an ORDER BY / HAVING name: a select alias, then a column
        string Ref(E e)
        {
            if (e is Num pos && pos.V == Math.Floor(pos.V) && pos.V >= 1 && pos.V <= output.Count) return output[(int)pos.V - 1].Internal;
            if (e is Agg a) return AggCol(a, null);
            if (e is Name n && n.Table == null) { var o = output.FindIndex(x => x.Shown.Equals(n.Col, StringComparison.OrdinalIgnoreCase)); if (o >= 0) return output[o].Internal; }
            if (e is Name m) return Col(m);
            if (grouped) throw new UserError("ORDER BY and HAVING take a column, an alias or an aggregate.");
            return Computed(e);
        }
        if (q.Where != null)
        {
            if (CondAgg(q.Where)) throw new UserError("WHERE cannot use SUM, COUNT, AVG, MIN or MAX; use HAVING.");
            spec.Where = JsonSerializer.SerializeToElement(Pred(q.Where, n => Col(n)));
        }
        if (q.Having != null) spec.Having = JsonSerializer.SerializeToElement(Pred(q.Having, n => Ref(n), e => Ref(e)));
        foreach (var o in q.Order) spec.Sort!.Add(new SortSpec(Ref(o.Expr), o.Desc));
        spec.Limit = q.Limit;
        if (spec.Computed!.Count == 0) spec.Computed = null; if (spec.Aggregates!.Count == 0) spec.Aggregates = null; if (spec.GroupBy!.Count == 0) spec.GroupBy = null; if (spec.Sort!.Count == 0) spec.Sort = null;
        spec.Columns = star && q.Items.Count == 1 ? null : output.Select(x => x.Internal).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        return new SqlPlan(spec, output, q.Limit);
    }

    static object Pred(C c, Func<Name, string> col, Func<E, string>? other = null)
    {
        string Side(E e) => e is Name n ? col(n) : other != null ? other(e) : throw new UserError("WHERE compares a column with a value or with another column.");
        bool IsCol(E e) => e is Name || (other != null && e is Agg);
        object? Lit(E e) => e switch { Num n => n.V, Str s => s.V, Bool b => b.V, Neg { V: Num n } => -n.V, Null => throw new UserError("Use IS NULL or IS NOT NULL to test for empty values."), _ => throw new UserError("Compare with a number, a quoted text or a column.") };
        string OpName(string op) => op switch { "=" => "eq", "<>" => "ne", "<" => "lt", "<=" => "le", ">" => "gt", ">=" => "ge", _ => throw new UserError("Unknown comparison.") };
        string Flip(string op) => op switch { "<" => ">", ">" => "<", "<=" => ">=", ">=" => "<=", _ => op };
        switch (c)
        {
            case And a: return new { op = "and", items = new[] { Pred(a.L, col, other), Pred(a.R, col, other) } };
            case Or o: return new { op = "or", items = new[] { Pred(o.L, col, other), Pred(o.R, col, other) } };
            case Not n: return new { op = "not", item = Pred(n.V, col, other) };
            case IsNull n: return new { op = "isNull", col = Side(n.L) };
            case In i: return new { op = "in", col = Side(i.L), values = i.Values.Select(Lit).ToArray() };
            case Between b: return new { op = "between", col = Side(b.L), a = Lit(b.A), b = Lit(b.B) };
            case Like l:
            {
                string pat = l.Pattern; if (pat.Contains('_')) throw new UserError("LIKE supports % only (text%, %text or %text%).");
                bool s = pat.StartsWith('%'), e = pat.EndsWith('%') && pat.Length > 1; string core = pat.Trim('%');
                if (core.Contains('%')) throw new UserError("LIKE supports % at the start, the end or both only (text%, %text or %text%).");
                string op = s && e ? "contains" : s ? "ends" : e ? "begins" : "eq";
                return new { op, col = Side(l.L), value = core };
            }
            case Cmp m:
                if (IsCol(m.L) && IsCol(m.R)) return new { op = "colCompare", col = Side(m.L), other = Side(m.R), cmp = OpName(m.Op) };
                if (IsCol(m.L)) return new { op = OpName(m.Op), col = Side(m.L), value = Lit(m.R) };
                if (IsCol(m.R)) return new { op = OpName(Flip(m.Op)), col = Side(m.R), value = Lit(m.L) };
                throw new UserError("A comparison needs a column on at least one side.");
        }
        throw new UserError("That condition is not supported.");
    }
}
