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

/// <summary>The edits over the open workbook: values, styles, column widths and freeze.
/// One user action is one undo record. Reading checks this before the file's own value.</summary>
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
    }

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
    public bool Dirty => _edits.Count > 0 || _colWidths.Count > 0 || _freeze.Count > 0;
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
    }
}
