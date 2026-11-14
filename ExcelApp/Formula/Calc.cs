// Live formula results for the editor: formulas the user typed are calculated over the fast copy plus the edits,
// the way Excel would show them. Results are kept until the edits change (EditSet.Version).
using System.Text.RegularExpressions;
using ExcelViewer.Core;

namespace ExcelApp.Formula;

public sealed class Calc : IFormulaGrid, IDisposable
{
    /// <summary>Evaluation nesting allowed before the engine switches to a top-to-bottom pass (keeps the stack small).</summary>
    const int MaxDepth = 300;

    readonly CacheData cache; readonly EditSet edits; readonly object editLock;
    readonly object gate = new();
    readonly Dictionary<int, SheetReader> readers = new();
    readonly Dictionary<(int Sheet, int Row, int Col), FValue> memo = new();
    readonly HashSet<(int Sheet, int Row, int Col)> active = new();
    readonly Dictionary<(int Sheet, int Row, int Col), (string Text, Node? Tree)> parsed = new();
    readonly Dictionary<string, int> sheetNames = new(StringComparer.OrdinalIgnoreCase);
    long version = long.MinValue; int depth;

    sealed class TooDeep : Exception { }

    public Calc(CacheData cache, EditSet edits, object editLock)
    {
        this.cache = cache; this.edits = edits; this.editLock = editLock;
        for (int i = 0; i < cache.Sheets.Count; i++) sheetNames.TryAdd(cache.Sheets[i].Name, i);
    }

    /// <summary>Lets go of the cache file so it can be replaced (the readers reopen on the next calculation).</summary>
    public void Release() { lock (gate) { foreach (var r in readers.Values) r.Dispose(); readers.Clear(); } }
    public void Dispose() => Release();

    // ---- IFormulaGrid: A1 coordinates (0-based row and column of the worksheet) -------------------
    public int SheetIndex(string name) => sheetNames.TryGetValue(name, out int i) ? i : -1;

    public (int Rows, int Cols) Extent(int sheet)
    {
        var d = cache.Sheets[sheet]; int rows, cols;
        lock (editLock) { rows = edits.RowCount(sheet, d.Rows); cols = edits.ColCount(sheet, d.Cols); }
        return (d.FirstRow - 1 + rows, d.FirstCol - 1 + cols);
    }

    public FValue Get(int sheet, int row, int col)
    {
        if (sheet < 0 || sheet >= cache.Sheets.Count) return FValue.Error("#REF!");
        var d = cache.Sheets[sheet]; int r = row + 1 - d.FirstRow, c = col + 1 - d.FirstCol; int rows, cols;
        lock (editLock) { rows = edits.RowCount(sheet, d.Rows); cols = edits.ColCount(sheet, d.Cols); }
        if (r < 0 || c < 0 || r >= rows || c >= cols) return FValue.Blank;
        return Value(sheet, r, c);
    }

    // ---- cache coordinates (row and column inside the copy) ---------------------------------------
    FValue Value(int sheet, int r, int c)
    {
        CellEdit? e; lock (editLock) edits.TryGet(sheet, r, c, out e!);
        if (e != null && !EditMarks.IsStyleOnly(e))
            return e.Kind switch
            {
                EditKind.Clear => FValue.Blank,
                EditKind.Number => FValue.Num(e.Num),
                EditKind.Text => FValue.Str(e.Text ?? ""),
                EditKind.Bool => FValue.Bool(e.Bool ?? e.Num != 0),
                _ => Compute(sheet, r, c, e.Formula ?? ""),
            };
        int pr, pc; lock (editLock) { pr = edits.PhysRow(sheet, r); pc = edits.PhysCol(sheet, c); }
        if (pr < 0 || pc < 0) return FValue.Blank;   // an inserted row or column
        if (!readers.TryGetValue(sheet, out var rd)) readers[sheet] = rd = new SheetReader(cache, cache.Sheets[sheet]);
        var col = rd.Column(pr, pc); int i = pr % rd.Sheet.BlockRows;
        return col.KindAt(i) switch
        {
            CellKind.Number => FValue.Num(col.Nums[i]),
            CellKind.Text => FValue.Str(col.TextAt(i)),
            CellKind.Bool => FValue.Bool(col.Nums[i] != 0),
            CellKind.Error => FValue.Error(SheetReader.ErrorText(col.Nums[i])),
            _ => FValue.Blank,
        };
    }

    void Fresh()
    {
        long v; lock (editLock) v = edits.Version;
        if (v == version) return;
        memo.Clear(); version = v;
        if (parsed.Count > 50_000) parsed.Clear();
    }

    /// <summary>The result of the formula in a cell (cache coordinates). Blank results show as 0, as in Excel.</summary>
    public FValue Result(int sheet, int r, int c, string formula)
    {
        lock (gate)
        {
            Fresh();
            try { return Compute(sheet, r, c, formula); }
            catch (TooDeep)
            {
                // A long chain (each cell using the one above): calculate every formula from the top so each step is short.
                active.Clear(); depth = 0;
                List<(int Sheet, int Row, int Col, string F)> all;
                lock (editLock) all = edits.All.Where(x => x.Value.Kind == EditKind.Formula).Select(x => (x.Key.Sheet, x.Key.Row, x.Key.Col, x.Value.Formula ?? "")).OrderBy(x => x.Sheet).ThenBy(x => x.Row).ThenBy(x => x.Col).ToList();
                for (int pass = 0; pass < 50; pass++)
                {
                    bool deep = false;
                    foreach (var f in all)
                    {
                        if (memo.ContainsKey((f.Sheet, f.Row, f.Col))) continue;
                        try { Compute(f.Sheet, f.Row, f.Col, f.F); } catch (TooDeep) { active.Clear(); depth = 0; deep = true; }
                    }
                    if (!deep) break;
                }
                try { return Compute(sheet, r, c, formula); } catch (TooDeep) { active.Clear(); depth = 0; return FValue.Error("#NUM!"); }
            }
        }
    }

    FValue Compute(int sheet, int r, int c, string formula)
    {
        var key = (sheet, r, c);
        if (memo.TryGetValue(key, out var known)) return known;
        if (!active.Add(key)) return FValue.Num(0);   // a circular reference: Excel shows 0 and warns
        if (++depth > MaxDepth) { depth--; active.Remove(key); throw new TooDeep(); }
        try
        {
            var d = cache.Sheets[sheet];
            var tree = Tree(key, formula);
            FValue v;
            if (tree == null) v = FValue.Error("#NAME?");
            else
            {
                var ev = new Evaluator(this, sheet) { Row = d.FirstRow - 1 + r, Col = d.FirstCol - 1 + c };
                try { v = ev.Evaluate(tree); }
                catch (TooDeep) { throw; }
                catch (Exception ex) when (ex is FormatException or ArgumentException or OverflowException or InvalidOperationException or IndexOutOfRangeException) { v = FValue.Error("#VALUE!"); }
            }
            if (v.Kind == ValueKind.Unsupported) v = FValue.Error("#NAME?");
            memo[key] = v;
            return v;
        }
        finally { depth--; active.Remove(key); }
    }

    Node? Tree((int, int, int) key, string formula)
    {
        if (parsed.TryGetValue(key, out var p) && p.Text == formula) return p.Tree;
        Node? tree; try { tree = new Parser(formula).Parse(); } catch (FormatException) { tree = null; }
        parsed[key] = (formula, tree);
        return tree;
    }

    /// <summary>Why Excel would refuse this formula as typed, or null when it is acceptable.</summary>
    public static string? Problem(string formula)
    {
        int depth = 0; bool inString = false, inSheet = false;
        foreach (char ch in formula)
        {
            if (inString) { if (ch == '"') inString = false; continue; }
            if (inSheet) { if (ch == '\'') inSheet = false; continue; }
            if (ch == '"') inString = true;
            else if (ch == '\'') inSheet = true;
            else if (ch == '(') depth++;
            else if (ch == ')' && --depth < 0) return "There's a problem with this formula: a closing bracket has no opening bracket.";
        }
        if (inString) return "There's a problem with this formula: a text value is missing its closing quote.";
        if (depth > 0) return "There's a problem with this formula: a bracket is not closed.";
        if (formula.Trim() == "=") return "There's a problem with this formula.";
        return null;
    }

    /// <summary>Adds the closing brackets Excel adds for you ("=SUM(A1:A3" becomes "=SUM(A1:A3)").</summary>
    public static string Complete(string formula)
    {
        int depth = 0; bool inString = false, inSheet = false;
        foreach (char ch in formula)
        {
            if (inString) { if (ch == '"') inString = false; continue; }
            if (inSheet) { if (ch == '\'') inSheet = false; continue; }
            if (ch == '"') inString = true; else if (ch == '\'') inSheet = true; else if (ch == '(') depth++; else if (ch == ')') depth--;
        }
        return depth > 0 && !inString && !inSheet ? formula + new string(')', depth) : formula;
    }
}

/// <summary>Moves the relative references of a formula, as Excel does when a formula is filled or pasted elsewhere.</summary>
public static class FormulaShift
{
    static readonly Regex CellRef = new(@"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$", RegexOptions.CultureInvariant);
    static readonly Regex ColRef = new(@"^(\$?)([A-Za-z]{1,3})$", RegexOptions.CultureInvariant);

    public static string Shift(string formula, int dRow, int dCol)
    {
        if (dRow == 0 && dCol == 0) return formula;
        var sb = new System.Text.StringBuilder(formula.Length + 8); int i = 0;
        while (i < formula.Length)
        {
            char ch = formula[i];
            if (ch == '"' || ch == '\'')
            {
                int start = i++; while (i < formula.Length) { if (formula[i] == ch) { if (i + 1 < formula.Length && formula[i + 1] == ch) { i += 2; continue; } i++; break; } i++; }
                sb.Append(formula, start, i - start); continue;
            }
            if (char.IsLetter(ch) || ch == '$' || ch == '_')
            {
                int start = i; while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '$' or '_' or '.')) i++;
                string word = formula[start..i];
                int k = i; while (k < formula.Length && formula[k] == ' ') k++;
                bool call = k < formula.Length && formula[k] == '(';
                bool sheetName = k < formula.Length && formula[k] == '!';
                char prev = start > 0 ? formula[start - 1] : '\0';
                if (call || sheetName) { sb.Append(word); continue; }
                var m = CellRef.Match(word);
                if (m.Success) { sb.Append(MoveCell(m, dRow, dCol)); continue; }
                var cm = ColRef.Match(word);
                if (cm.Success && (prev == ':' || (k < formula.Length && formula[k] == ':')) && !word.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
                {
                    if (cm.Groups[1].Value == "$") { sb.Append(word); continue; }
                    int c = Letters(cm.Groups[2].Value) + dCol;
                    sb.Append(c < 1 || c > 16384 ? "#REF!" : ExcelViewer.Core.CacheBuilder.Letters(c));
                    continue;
                }
                sb.Append(word); continue;
            }
            sb.Append(ch); i++;
        }
        return sb.ToString();
    }

    static string MoveCell(Match m, int dRow, int dCol)
    {
        bool absCol = m.Groups[1].Value == "$", absRow = m.Groups[3].Value == "$";
        int col = Letters(m.Groups[2].Value) + (absCol ? 0 : dCol);
        long row = long.Parse(m.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture) + (absRow ? 0 : dRow);
        if (col < 1 || col > 16384 || row < 1 || row > 1048576) return "#REF!";
        return (absCol ? "$" : "") + ExcelViewer.Core.CacheBuilder.Letters(col) + (absRow ? "$" : "") + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    static int Letters(string s) { int n = 0; foreach (char c in s.ToUpperInvariant()) n = n * 26 + c - 'A' + 1; return n; }

    static readonly Regex RowRef = new(@"^(\$?)(\d{1,7})$", RegexOptions.CultureInvariant);

    /// <summary>Moves a formula's references as Excel does when rows (or columns) are inserted or deleted on a sheet.
    /// <paramref name="at"/> is the first Excel row or column number affected (1-based). References at or after it move by
    /// <paramref name="count"/>; a range that spans it grows on insert and shrinks on delete; a reference wholly inside
    /// deleted rows becomes #REF!. <paramref name="unqualifiedIsTarget"/>: the formula sits on the changed sheet, so
    /// references without a sheet name point at it. Absolute ($) references move too, as in Excel.</summary>
    public static string Structural(string formula, bool unqualifiedIsTarget, string targetSheet, bool rows, int at, int count, bool insert)
    {
        if (count <= 0 || string.IsNullOrEmpty(formula)) return formula;
        int last = at + count - 1;
        int Move(int v) => insert ? (v >= at ? v + count : v) : (v < at ? v : v > last ? v - count : -1);
        (int, int)? Span(int a, int b)
        {
            if (insert) return a >= at ? (a + count, b + count) : b >= at ? (a, b + count) : (a, b);
            if (a >= at && b <= last) return null;
            int na = a < at ? a : a > last ? a - count : at, nb = b < at ? b : b > last ? b - count : at - 1;
            return nb < na ? null : (na, nb);
        }
        var sb = new System.Text.StringBuilder(formula.Length + 8); int i = 0;
        while (i < formula.Length)
        {
            char ch = formula[i];
            if (ch == '"')
            {
                int start = i++; while (i < formula.Length) { if (formula[i] == '"') { if (i + 1 < formula.Length && formula[i + 1] == '"') { i += 2; continue; } i++; break; } i++; }
                sb.Append(formula, start, i - start); continue;
            }
            // an optional sheet name: 'My sheet'! or Sheet1!
            string? sheet = null; int refStart = i;
            if (ch == '\'')
            {
                int j = i + 1; var name = new System.Text.StringBuilder();
                while (j < formula.Length) { if (formula[j] == '\'') { if (j + 1 < formula.Length && formula[j + 1] == '\'') { name.Append('\''); j += 2; continue; } j++; break; } name.Append(formula[j++]); }
                if (j < formula.Length && formula[j] == '!') { sheet = name.ToString(); i = j + 1; }
                else { sb.Append(formula, i, j - i); i = j; continue; }
            }
            else if (char.IsLetter(ch) || ch == '_')
            {
                int j = i; while (j < formula.Length && (char.IsLetterOrDigit(formula[j]) || formula[j] is '_' or '.')) j++;
                if (j < formula.Length && formula[j] == '!') { sheet = formula[i..j]; i = j + 1; }
            }
            // the reference itself: A1, $A$1, A1:B2, A:C, 3:5
            int r0 = i; while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '$' or '_' or '.')) i++;
            string first = formula[r0..i];
            if (first.Length == 0) { if (sheet != null) sb.Append(formula, refStart, i - refStart); else { sb.Append(ch); i++; } continue; }
            int k = i; while (k < formula.Length && formula[k] == ' ') k++;
            if (k < formula.Length && formula[k] == '(') { sb.Append(formula, refStart, i - refStart); continue; }   // a function name
            string? second = null; int end = i;
            if (i < formula.Length && formula[i] == ':')
            {
                int j = i + 1; while (j < formula.Length && (char.IsLetterOrDigit(formula[j]) || formula[j] == '$')) j++;
                if (j > i + 1) { second = formula[(i + 1)..j]; end = j; }
            }
            char before = refStart > 0 ? formula[refStart - 1] : '\0';
            bool inName = char.IsLetterOrDigit(before) || before is '_' or '.';
            bool target = sheet == null ? unqualifiedIsTarget : string.Equals(sheet, targetSheet, StringComparison.OrdinalIgnoreCase);
            string prefix = formula[refStart..r0];
            string? moved = inName || !target ? null : MoveRef(first, second, rows, Move, Span);
            if (moved == null) { sb.Append(formula, refStart, end - refStart); i = end; continue; }
            sb.Append(prefix).Append(moved); i = end;
        }
        return sb.ToString();
    }

    /// <summary>The moved reference text, "#REF!" when it was deleted, or null when the text is not a reference.</summary>
    static string? MoveRef(string a, string? b, bool rows, Func<int, int> move, Func<int, int, (int, int)?> span)
    {
        var ca = CellRef.Match(a);
        if (b == null)
        {
            if (!ca.Success || a.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return null;
            int col = Letters(ca.Groups[2].Value), row = int.Parse(ca.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture);
            int v = move(rows ? row : col); if (v < 1 || v > (rows ? 1048576 : 16384)) return "#REF!";
            return Cell(ca.Groups[1].Value, rows ? col : v, ca.Groups[3].Value, rows ? v : row);
        }
        var cb = CellRef.Match(b);
        if (ca.Success && cb.Success)
        {
            int c1 = Letters(ca.Groups[2].Value), c2 = Letters(cb.Groups[2].Value);
            int r1 = int.Parse(ca.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture), r2 = int.Parse(cb.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture);
            var s = rows ? span(Math.Min(r1, r2), Math.Max(r1, r2)) : span(Math.Min(c1, c2), Math.Max(c1, c2));
            if (s == null) return "#REF!";
            var (x1, x2) = s.Value; if (x2 > (rows ? 1048576 : 16384)) return "#REF!";
            return rows ? Cell(ca.Groups[1].Value, c1, ca.Groups[3].Value, x1) + ":" + Cell(cb.Groups[1].Value, c2, cb.Groups[3].Value, x2)
                        : Cell(ca.Groups[1].Value, x1, ca.Groups[3].Value, r1) + ":" + Cell(cb.Groups[1].Value, x2, cb.Groups[3].Value, r2);
        }
        var la = ColRef.Match(a); var lb = ColRef.Match(b);
        if (la.Success && lb.Success)   // whole columns, such as A:C
        {
            if (rows) return a + ":" + b;
            var s = span(Math.Min(Letters(la.Groups[2].Value), Letters(lb.Groups[2].Value)), Math.Max(Letters(la.Groups[2].Value), Letters(lb.Groups[2].Value)));
            if (s == null || s.Value.Item2 > 16384) return "#REF!";
            return la.Groups[1].Value + ExcelViewer.Core.CacheBuilder.Letters(s.Value.Item1) + ":" + lb.Groups[1].Value + ExcelViewer.Core.CacheBuilder.Letters(s.Value.Item2);
        }
        var ra = RowRef.Match(a); var rb = RowRef.Match(b);
        if (ra.Success && rb.Success)   // whole rows, such as 3:5
        {
            if (!rows) return a + ":" + b;
            int x = int.Parse(ra.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture), y = int.Parse(rb.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
            var s = span(Math.Min(x, y), Math.Max(x, y));
            if (s == null || s.Value.Item2 > 1048576) return "#REF!";
            return ra.Groups[1].Value + s.Value.Item1.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + rb.Groups[1].Value + s.Value.Item2.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        return null;
    }
    static string Cell(string absCol, int col, string absRow, int row) => absCol + ExcelViewer.Core.CacheBuilder.Letters(col) + absRow + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
