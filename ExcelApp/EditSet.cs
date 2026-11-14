// Value, style, width and freeze edits with undo/redo, and the typed-input parser of plan 2.4.
// Standard library only: this file has no dependency on ExcelViewer or Excel.
using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelApp;

/// <summary>What a cell edit holds. Clear returns the cell to the file's own value.</summary>
public enum EditKind { Clear, Number, Text, Bool, Formula }

/// <summary>A change to one cell's look. Nulls mean "not touched by this delta".</summary>
public sealed record StyleDelta(
    bool? Bold = null, bool? Italic = null, bool? Underline = null,
    int? FontColor = null, int? Fill = null, int? Align = null,
    string? NumberFormat = null, string? Border = null)
{
    /// <summary>A copy where every non-null field of <paramref name="over"/> wins over this one.</summary>
    public StyleDelta Merge(StyleDelta? over)
    {
        if (over == null) return this;
        return new StyleDelta(
            Bold: over.Bold ?? Bold,
            Italic: over.Italic ?? Italic,
            Underline: over.Underline ?? Underline,
            FontColor: over.FontColor ?? FontColor,
            Fill: over.Fill ?? Fill,
            Align: over.Align ?? Align,
            NumberFormat: over.NumberFormat ?? NumberFormat,
            Border: over.Border ?? Border);
    }
}

/// <summary>One cell's edited value and optional style.</summary>
public sealed record CellEdit(
    EditKind Kind,
    double Num = 0,
    string? Text = null,
    string? Formula = null,
    bool? Bool = null,
    StyleDelta? Style = null);

/// <summary>Turns what the user typed into a <see cref="CellEdit"/>, following plan 2.4.</summary>
public static class InputParser
{
    const string General = "General";

    /// <summary>Parses one typed value in the given culture.
    /// <paramref name="currentFormat"/> is "General" for an unformatted cell.
    /// <paramref name="impliedFormat"/> is "0.00%" or "yyyy-mm-dd" when the plan says to imply one, else null.</summary>
    public static CellEdit Parse(string input, CultureInfo culture, string currentFormat, bool date1904, out string? impliedFormat)
    {
        impliedFormat = null;
        input ??= "";
        culture ??= CultureInfo.CurrentCulture;
        bool general = string.IsNullOrEmpty(currentFormat) ||
                       string.Equals(currentFormat, General, StringComparison.OrdinalIgnoreCase);

        // A leading apostrophe means "this is text": the only way to type text that looks like a formula.
        if (input.Length > 0 && input[0] == '\'')
            return new CellEdit(EditKind.Text, Text: input.Substring(1));

        string trimmed = input.Trim();
        if (trimmed.Length == 0)
            return new CellEdit(EditKind.Clear);

        // "=..." is a formula. Text that starts with "=" is reachable only through the apostrophe above.
        if (trimmed[0] == '=')
            return new CellEdit(EditKind.Formula, Formula: input);

        // TRUE / FALSE in any case.
        if (string.Equals(trimmed, "TRUE", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(trimmed, "FALSE", StringComparison.OrdinalIgnoreCase))
        {
            bool b = trimmed[0] == 't' || trimmed[0] == 'T';
            return new CellEdit(EditKind.Bool, Num: b ? 1 : 0, Bool: b);
        }

        // A number, in the user's culture, with any allowed signs, separators and symbols.
        if (double.TryParse(trimmed, NumberStyles.Any, culture, out double num))
            return new CellEdit(EditKind.Number, Num: num);

        // A percent such as "12.5%" is the number scaled by 100, with an implied percent format when General.
        if (trimmed.Length > 1 && trimmed[trimmed.Length - 1] == '%')
        {
            string body = trimmed.Substring(0, trimmed.Length - 1).Trim();
            if (double.TryParse(body, NumberStyles.Any, culture, out double pct))
            {
                StyleDelta? style = general ? new StyleDelta(NumberFormat: "0.00%") : null;
                if (style != null) impliedFormat = "0.00%";
                return new CellEdit(EditKind.Number, Num: pct / 100.0, Style: style);
            }
        }

        // A time on its own ("13:45", "9:30 PM", "8:05:10") is a fraction of a day, as Excel stores it, with a time format when General.
        var tm = System.Text.RegularExpressions.Regex.Match(trimmed, @"^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$");
        if (tm.Success)
        {
            int h = int.Parse(tm.Groups[1].Value, CultureInfo.InvariantCulture), m = int.Parse(tm.Groups[2].Value, CultureInfo.InvariantCulture);
            int s = tm.Groups[3].Success ? int.Parse(tm.Groups[3].Value, CultureInfo.InvariantCulture) : 0;
            bool ampm = tm.Groups[4].Success;
            if (ampm) { if (h < 1 || h > 12) h = -1; else h = h % 12 + (char.ToUpperInvariant(tm.Groups[4].Value[0]) == 'P' ? 12 : 0); }
            if (h >= 0 && h < 24 && m < 60 && s < 60)
            {
                string code = ampm ? (tm.Groups[3].Success ? "h:mm:ss AM/PM" : "h:mm AM/PM") : (tm.Groups[3].Success ? "h:mm:ss" : "h:mm");
                StyleDelta? style = general ? new StyleDelta(NumberFormat: code) : null;
                if (style != null) impliedFormat = code;
                return new CellEdit(EditKind.Number, Num: (h * 3600 + m * 60 + s) / 86400.0, Style: style);
            }
        }

        // A date the culture accepts becomes an Excel serial, with an implied date format when General.
        if (DateTime.TryParse(trimmed, culture, DateTimeStyles.None, out DateTime dt))
        {
            DateTime epoch = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
            double serial = (dt - epoch).TotalDays;
            StyleDelta? style = general ? new StyleDelta(NumberFormat: "yyyy-mm-dd") : null;
            if (style != null) impliedFormat = "yyyy-mm-dd";
            return new CellEdit(EditKind.Number, Num: serial, Style: style);
        }

        // Everything else is text, kept exactly as typed.
        return new CellEdit(EditKind.Text, Text: input);
    }
}

/// <summary>A chart over a range of a sheet (logical rows and columns, inclusive). X, Y, W, H: its card in the window, in pixels.
/// Type: column, bar, line, area, pie or scatter.</summary>
public sealed record ChartSpec(string Id, int Sheet, string Type, string? Title, int R1, int C1, int R2, int C2, double X, double Y, double W, double H);

/// <summary>Rows (or columns) inserted or deleted on a sheet. At is the first logical row or column (0-based, as the editor shows them).</summary>
public sealed record StructOp(int Sheet, bool Rows, int At, int Count, bool Insert);

/// <summary>The edits over the open workbook: values, styles, column widths, freeze, and inserted or deleted rows and columns.
/// One user action is one undo record. Reading checks this before the file's own value.
/// Positions are logical (what the editor shows); a sheet's layout maps them to the fast copy's rows and columns.</summary>
public sealed class EditSet
{
    readonly Dictionary<(int Sheet, int Row, int Col), CellEdit> _edits = new();
    readonly Dictionary<(int Sheet, int Col), double> _colWidths = new();
    readonly Dictionary<int, (int Rows, int Cols)> _freeze = new();
    readonly List<Change> _undo = new();
    readonly List<Change> _redo = new();

    sealed class Change
    {
        public int Sheet;
        public readonly List<(int Row, int Col, CellEdit? Before, CellEdit? After)> Cells = new();
        // An insert or delete: the whole state it replaced, so undo puts back exactly what was there.
        public StructOp? Op; public int PhysRows, PhysCols;
        public List<int>? RowsBefore, ColsBefore;
        public Dictionary<(int Sheet, int Row, int Col), CellEdit>? EditsBefore;
        public Dictionary<(int Sheet, int Col), double>? WidthsBefore;
        public Dictionary<string, ChartSpec>? ChartsBefore;
        // A chart added, changed or removed.
        public bool IsChart; public ChartSpec? ChartBefore, ChartAfter;
    }

    readonly Dictionary<string, ChartSpec> _charts = new();
    /// <summary>The workbook's charts, made in this editor (saved into the file through Excel).</summary>
    public IReadOnlyCollection<ChartSpec> Charts => _charts.Values;
    public ChartSpec? Chart(string id) => _charts.TryGetValue(id, out var c) ? c : null;
    /// <summary>Adds (before null), changes, or removes (after null) a chart as one undo step.</summary>
    public void SetChart(ChartSpec? before, ChartSpec? after)
    {
        if (before == null && after == null) return;
        PutChart(before, after);
        _undo.Add(new Change { Sheet = (after ?? before)!.Sheet, IsChart = true, ChartBefore = before, ChartAfter = after }); _redo.Clear();
    }
    void PutChart(ChartSpec? remove, ChartSpec? add) { if (remove != null) _charts.Remove(remove.Id); if (add != null) _charts[add.Id] = add; Version++; }

    // ---- layout: logical row/column -> fast-copy row/column (-1 = inserted, blank). Null = unchanged.
    readonly Dictionary<int, List<int>> _rowsOf = new(), _colsOf = new();
    readonly Dictionary<int, int[]> _rowInverse = new(), _colInverse = new();
    readonly List<StructOp> _ops = new();

    /// <summary>Moves a formula's references for an insert or delete: (formula's sheet, formula, op) -> new formula. Set by the session.</summary>
    public Func<int, string, StructOp, string>? ShiftFormula { get; set; }

    /// <summary>The inserts and deletes in effect, in the order they were made (the save replays them in Excel).</summary>
    public IReadOnlyList<StructOp> Ops => _ops;
    public bool HasLayout(int sheet) => _rowsOf.ContainsKey(sheet) || _colsOf.ContainsKey(sheet);
    public int RowCount(int sheet, int physRows) => _rowsOf.TryGetValue(sheet, out var l) ? l.Count : physRows;
    public int ColCount(int sheet, int physCols) => _colsOf.TryGetValue(sheet, out var l) ? l.Count : physCols;
    /// <summary>The fast-copy row of a logical row; -1 for an inserted (blank) row or one past the end.</summary>
    public int PhysRow(int sheet, int row) => _rowsOf.TryGetValue(sheet, out var l) ? (row >= 0 && row < l.Count ? l[row] : -1) : row;
    public int PhysCol(int sheet, int col) => _colsOf.TryGetValue(sheet, out var l) ? (col >= 0 && col < l.Count ? l[col] : -1) : col;
    /// <summary>The logical row showing a fast-copy row; -1 when that row was deleted.</summary>
    public int LogicalRow(int sheet, int phys) => Inverse(_rowsOf, _rowInverse, sheet, phys);
    public int LogicalCol(int sheet, int phys) => Inverse(_colsOf, _colInverse, sheet, phys);
    static int Inverse(Dictionary<int, List<int>> map, Dictionary<int, int[]> cache, int sheet, int phys)
    {
        if (!map.TryGetValue(sheet, out var l)) return phys;
        if (!cache.TryGetValue(sheet, out var inv))
        {
            int max = -1; foreach (int p in l) if (p > max) max = p;
            inv = new int[max + 1]; Array.Fill(inv, -1);
            for (int i = 0; i < l.Count; i++) if (l[i] >= 0) inv[l[i]] = i;
            cache[sheet] = inv;
        }
        return phys >= 0 && phys < inv.Length ? inv[phys] : -1;
    }

    /// <summary>Inserts or deletes rows or columns as one undo step. Edits below move, edits in deleted rows go, and every
    /// formula's references move as Excel moves them. <paramref name="physRows"/>/<paramref name="physCols"/> are the fast copy's size.</summary>
    public void Structure(StructOp op, int physRows, int physCols)
    {
        int size = op.Rows ? RowCount(op.Sheet, physRows) : ColCount(op.Sheet, physCols);
        if (op.Count < 1) throw new ArgumentOutOfRangeException(nameof(op));
        if (op.Insert ? op.At < 0 || op.At > size : op.At < 0 || op.At + op.Count > size) throw new ArgumentOutOfRangeException(nameof(op));
        var change = new Change
        {
            Sheet = op.Sheet, Op = op, PhysRows = physRows, PhysCols = physCols,
            RowsBefore = _rowsOf.TryGetValue(op.Sheet, out var rb) ? new List<int>(rb) : null,
            ColsBefore = _colsOf.TryGetValue(op.Sheet, out var cb) ? new List<int>(cb) : null,
            EditsBefore = new Dictionary<(int, int, int), CellEdit>(_edits),
            WidthsBefore = _colWidths.Where(kv => kv.Key.Sheet == op.Sheet).ToDictionary(kv => kv.Key, kv => kv.Value),
            ChartsBefore = new Dictionary<string, ChartSpec>(_charts),
        };
        Run(op, physRows, physCols);
        _undo.Add(change); _redo.Clear();
    }

    void Run(StructOp op, int physRows, int physCols)
    {
        int last = op.At + op.Count - 1;
        int Move(int v) => op.Insert ? (v >= op.At ? v + op.Count : v) : (v < op.At ? v : v > last ? v - op.Count : -1);
        var moved = new List<KeyValuePair<(int, int, int), CellEdit>>(_edits.Count);
        foreach (var kv in _edits)
        {
            var (sh, r, c) = kv.Key; var e = kv.Value;
            if (sh == op.Sheet) { if (op.Rows) r = Move(r); else c = Move(c); if (r < 0 || c < 0) continue; }
            if (e.Kind == EditKind.Formula && ShiftFormula != null && e.Formula != null) e = e with { Formula = ShiftFormula(sh, e.Formula, op) };
            moved.Add(new((sh, r, c), e));
        }
        _edits.Clear(); foreach (var kv in moved) _edits[kv.Key] = kv.Value;
        if (!op.Rows)
        {
            var widths = _colWidths.Where(kv => kv.Key.Sheet == op.Sheet).ToList();
            foreach (var kv in widths) _colWidths.Remove(kv.Key);
            foreach (var kv in widths) { int c = Move(kv.Key.Col); if (c >= 0) _colWidths[(op.Sheet, c)] = kv.Value; }
        }
        // charts on the sheet: their ranges move, grow and shrink as formula ranges do; a chart whose range was deleted goes
        foreach (var ch in _charts.Values.Where(x => x.Sheet == op.Sheet).ToList())
        {
            int a = op.Rows ? ch.R1 : ch.C1, b = op.Rows ? ch.R2 : ch.C2, na, nb;
            if (op.Insert) { na = a >= op.At ? a + op.Count : a; nb = b >= op.At ? b + op.Count : b; }
            else
            {
                if (a >= op.At && b <= last) { _charts.Remove(ch.Id); continue; }
                na = a < op.At ? a : a > last ? a - op.Count : op.At; nb = b < op.At ? b : b > last ? b - op.Count : op.At - 1;
            }
            _charts[ch.Id] = op.Rows ? ch with { R1 = na, R2 = nb } : ch with { C1 = na, C2 = nb };
        }
        var map = op.Rows ? _rowsOf : _colsOf;
        if (!map.TryGetValue(op.Sheet, out var l)) map[op.Sheet] = l = Enumerable.Range(0, op.Rows ? physRows : physCols).ToList();
        if (op.Insert) l.InsertRange(op.At, Enumerable.Repeat(-1, op.Count)); else l.RemoveRange(op.At, op.Count);
        (op.Rows ? _rowInverse : _colInverse).Remove(op.Sheet);
        _ops.Add(op);
        Recount();
    }

    void Recount() { _formulas = _edits.Values.Count(e => e.Kind == EditKind.Formula); Version++; }

    void Restore(Change c)
    {
        _edits.Clear(); foreach (var kv in c.EditsBefore!) _edits[kv.Key] = kv.Value;
        foreach (var k in _colWidths.Keys.Where(k => k.Sheet == c.Sheet).ToList()) _colWidths.Remove(k);
        foreach (var kv in c.WidthsBefore!) _colWidths[kv.Key] = kv.Value;
        if (c.RowsBefore != null) _rowsOf[c.Sheet] = new List<int>(c.RowsBefore); else _rowsOf.Remove(c.Sheet);
        if (c.ColsBefore != null) _colsOf[c.Sheet] = new List<int>(c.ColsBefore); else _colsOf.Remove(c.Sheet);
        _rowInverse.Remove(c.Sheet); _colInverse.Remove(c.Sheet);
        _charts.Clear(); foreach (var kv in c.ChartsBefore!) _charts[kv.Key] = kv.Value;
        _ops.RemoveAt(_ops.Count - 1);
        Recount();
    }

    /// <summary>True when the last undo or redo was an insert or delete (the interface reads the whole sheet again).</summary>
    public bool LastWasStructural { get; private set; }
    /// <summary>True when the last undo or redo was a chart change.</summary>
    public bool LastWasChart { get; private set; }

    /// <summary>Goes up on every change to the cell edits; formula results computed for an older version are stale.</summary>
    public long Version { get; private set; }

    /// <summary>True when any cell holds a formula edit.</summary>
    public bool HasFormulas => _formulas > 0;
    int _formulas;
    void Track(CellEdit? before, CellEdit? after)
    {
        if (before?.Kind == EditKind.Formula) _formulas--;
        if (after?.Kind == EditKind.Formula) _formulas++;
        Version++;
    }

    /// <summary>True when an edit, a width or a freeze change exists.</summary>
    public bool Dirty => _edits.Count > 0 || _colWidths.Count > 0 || _freeze.Count > 0 || _ops.Count > 0 || _charts.Count > 0;
    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;

    /// <summary>Looks up the edit for a cell; false means the file's own value shows.</summary>
    public bool TryGet(int sheet, int row, int col, out CellEdit edit)
        => _edits.TryGetValue((sheet, row, col), out edit!);

    /// <summary>Every edit, keyed by sheet, row and column.</summary>
    public IReadOnlyCollection<KeyValuePair<(int Sheet, int Row, int Col), CellEdit>> All => _edits;

    /// <summary>Applies many cells as one action. <c>after == null</c> removes the edit.
    /// Returns the cells whose value actually changed, in the given order.</summary>
    public List<(int Row, int Col)> Apply(int sheet, IReadOnlyList<(int Row, int Col, CellEdit? After)> cells)
    {
        var change = new Change { Sheet = sheet };
        var changed = new List<(int Row, int Col)>();
        foreach (var (row, col, after) in cells)
        {
            var key = (sheet, row, col);
            _edits.TryGetValue(key, out CellEdit? before);
            if (Equals(before, after)) continue; // nothing changes; do not add to the record
            change.Cells.Add((row, col, before, after));
            Track(before, after);
            if (after == null) _edits.Remove(key); else _edits[key] = after;
            changed.Add((row, col));
        }
        if (change.Cells.Count > 0)
        {
            _undo.Add(change);
            _redo.Clear();
        }
        return changed;
    }

    /// <summary>Reverses the last action. Sheet is -1 when there is nothing to undo.</summary>
    public (int Sheet, List<(int Row, int Col)> Changed) Undo()
    {
        if (_undo.Count == 0) return (-1, new List<(int Row, int Col)>());
        var change = _undo[_undo.Count - 1];
        _undo.RemoveAt(_undo.Count - 1);
        LastWasStructural = change.Op != null; LastWasChart = change.IsChart;
        if (change.IsChart) { PutChart(change.ChartAfter, change.ChartBefore); _redo.Add(change); return (change.Sheet, new List<(int Row, int Col)>()); }
        if (change.Op != null) { Restore(change); _redo.Add(change); return (change.Sheet, new List<(int Row, int Col)>()); }
        var changed = new List<(int Row, int Col)>(change.Cells.Count);
        for (int i = change.Cells.Count - 1; i >= 0; i--)
        {
            var (row, col, before, after) = change.Cells[i];
            var key = (change.Sheet, row, col);
            Track(after, before);
            if (before == null) _edits.Remove(key); else _edits[key] = before;
            changed.Add((row, col));
        }
        _redo.Add(change);
        return (change.Sheet, changed);
    }

    /// <summary>Re-applies the last undone action. Sheet is -1 when there is nothing to redo.</summary>
    public (int Sheet, List<(int Row, int Col)> Changed) Redo()
    {
        if (_redo.Count == 0) return (-1, new List<(int Row, int Col)>());
        var change = _redo[_redo.Count - 1];
        _redo.RemoveAt(_redo.Count - 1);
        LastWasStructural = change.Op != null; LastWasChart = change.IsChart;
        if (change.IsChart) { PutChart(change.ChartBefore, change.ChartAfter); _undo.Add(change); return (change.Sheet, new List<(int Row, int Col)>()); }
        if (change.Op != null) { Run(change.Op, change.PhysRows, change.PhysCols); _undo.Add(change); return (change.Sheet, new List<(int Row, int Col)>()); }
        var changed = new List<(int Row, int Col)>(change.Cells.Count);
        foreach (var (row, col, before, after) in change.Cells)
        {
            var key = (change.Sheet, row, col);
            Track(before, after);
            if (after == null) _edits.Remove(key); else _edits[key] = after;
            changed.Add((row, col));
        }
        _undo.Add(change);
        return (change.Sheet, changed);
    }

    public void SetColWidth(int sheet, int col, double px) => _colWidths[(sheet, col)] = px;
    public IReadOnlyDictionary<(int Sheet, int Col), double> ColWidths => _colWidths;

    public void SetFreeze(int sheet, int rows, int cols) => _freeze[sheet] = (rows, cols);
    public IReadOnlyDictionary<int, (int Rows, int Cols)> Freeze => _freeze;

    /// <summary>Forgets everything, as after a save.</summary>
    public void Clear()
    {
        _edits.Clear();
        _formulas = 0; Version++;
        _colWidths.Clear();
        _freeze.Clear();
        _undo.Clear();
        _redo.Clear();
        _rowsOf.Clear(); _colsOf.Clear(); _rowInverse.Clear(); _colInverse.Clear(); _ops.Clear(); _charts.Clear();
    }
}
