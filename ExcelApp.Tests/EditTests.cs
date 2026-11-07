// Lane B tests: typed-input parsing, undo/redo and the one-action-one-record rule. No Excel needed.
using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelApp;

/// <summary>Tests for the typed-input parser and the edit overlay. See EditSet.cs.</summary>
public static class EditTests
{
    /// <summary>Runs the checks. The thunk reports pass/fail: name, ok, detail.</summary>
    public static void Run(Action<string, bool, string> check)
    {
        ParsingTable(check);
        PercentAndDate(check);
        UndoRedo(check);
        OnePasteOneRecord(check);
    }

    static (CellEdit Edit, string? Implied) P(string input, string culture, string format, bool date1904 = false)
    {
        var edit = InputParser.Parse(input, new CultureInfo(culture), format, date1904, out var implied);
        return (edit, implied);
    }

    static void ParsingTable(Action<string, bool, string> check)
    {
        const string en = "en-US";
        const string de = "de-DE";

        var r = P("", en, "General");
        check("parse: empty is Clear", r.Edit.Kind == EditKind.Clear, r.Edit.Kind.ToString());

        r = P("apple", en, "General");
        check("parse: plain text", r.Edit.Kind == EditKind.Text && r.Edit.Text == "apple", r.Edit.Text ?? "null");

        r = P("=A1*2", en, "General");
        check("parse: formula", r.Edit.Kind == EditKind.Formula && r.Edit.Formula == "=A1*2", r.Edit.Formula ?? "null");

        r = P("TRUE", en, "General");
        check("parse: TRUE is bool true", r.Edit.Kind == EditKind.Bool && r.Edit.Bool == true && r.Edit.Num == 1, r.Edit.Bool?.ToString() ?? "null");

        r = P("false", en, "General");
        check("parse: false (any case) is bool false", r.Edit.Kind == EditKind.Bool && r.Edit.Bool == false, r.Edit.Bool?.ToString() ?? "null");

        r = P("12.5", en, "General");
        check("parse: en-US number 12.5", r.Edit.Kind == EditKind.Number && r.Edit.Num == 12.5, r.Edit.Num.ToString(CultureInfo.InvariantCulture));

        r = P("1,234.56", en, "General");
        check("parse: en-US thousands 1234.56", r.Edit.Kind == EditKind.Number && r.Edit.Num == 1234.56, r.Edit.Num.ToString(CultureInfo.InvariantCulture));

        r = P("'=x", en, "General");
        check("parse: apostrophe makes '=x' text", r.Edit.Kind == EditKind.Text && r.Edit.Text == "=x", r.Edit.Text ?? "null");

        r = P("'123", en, "General");
        check("parse: apostrophe makes '123' text", r.Edit.Kind == EditKind.Text && r.Edit.Text == "123", r.Edit.Text ?? "null");

        r = P("12,5", de, "General");
        check("parse: de-DE number 12,5 -> 12.5", r.Edit.Kind == EditKind.Number && r.Edit.Num == 12.5, r.Edit.Num.ToString(CultureInfo.InvariantCulture));

        r = P("1.234,56", de, "General");
        check("parse: de-DE 1.234,56 -> 1234.56", r.Edit.Kind == EditKind.Number && r.Edit.Num == 1234.56, r.Edit.Num.ToString(CultureInfo.InvariantCulture));
    }

    static void PercentAndDate(Action<string, bool, string> check)
    {
        const string en = "en-US";

        var r = P("12.5%", en, "General");
        check("percent: 12.5% is 0.125", r.Edit.Kind == EditKind.Number && r.Edit.Num == 0.125, r.Edit.Num.ToString(CultureInfo.InvariantCulture));
        check("percent: implied 0.00% on a General cell", r.Implied == "0.00%" && r.Edit.Style?.NumberFormat == "0.00%", r.Implied ?? "null");

        r = P("12.5%", en, "0.00%");
        check("percent: no implied format when the cell is already formatted", r.Edit.Num == 0.125 && r.Implied == null && r.Edit.Style == null, r.Implied ?? "null");

        r = P("2020-01-01", en, "General");
        check("date: 2020-01-01 is serial 43831 (1900 system)", r.Edit.Kind == EditKind.Number && r.Edit.Num == 43831, r.Edit.Num.ToString(CultureInfo.InvariantCulture));
        check("date: implied yyyy-mm-dd on a General cell", r.Implied == "yyyy-mm-dd" && r.Edit.Style?.NumberFormat == "yyyy-mm-dd", r.Implied ?? "null");

        r = P("2020-01-01", en, "General", date1904: true);
        check("date: 2020-01-01 is serial 42369 (1904 system)", r.Edit.Num == 42369, r.Edit.Num.ToString(CultureInfo.InvariantCulture));

        r = P("2020-01-01", en, "dd/mm/yyyy");
        check("date: no implied format when the cell is already formatted", r.Implied == null && r.Edit.Style == null, r.Implied ?? "null");

        r = P("12,5%", "de-DE", "General");
        check("percent: de-DE 12,5% is 0.125", r.Edit.Num == 0.125, r.Edit.Num.ToString(CultureInfo.InvariantCulture));
    }

    static void UndoRedo(Action<string, bool, string> check)
    {
        var set = new EditSet();
        var a = new CellEdit(EditKind.Number, Num: 5);
        var b = new CellEdit(EditKind.Text, Text: "x");

        var changed = set.Apply(0, new List<(int, int, CellEdit?)> { (1, 0, a) });
        check("apply: one cell reports one change", changed.Count == 1 && changed[0] == (1, 0), changed.Count.ToString());
        check("apply: dirty, can undo, cannot redo", set.Dirty && set.CanUndo && !set.CanRedo, "");
        check("apply: edit reads back", set.TryGet(0, 1, 0, out var g1) && Equals(g1, a), "");

        set.Apply(0, new List<(int, int, CellEdit?)> { (2, 0, b) });

        var (s1, _) = set.Undo();
        check("undo: reports sheet 0", s1 == 0, s1.ToString());
        check("undo: removes the second edit", !set.TryGet(0, 2, 0, out _), "");
        check("undo: keeps the first edit", set.TryGet(0, 1, 0, out var g2) && Equals(g2, a), "");

        var (s2, _) = set.Undo();
        check("undo: second undo empties the overlay", !set.TryGet(0, 1, 0, out _) && !set.Dirty, "");
        check("undo: nothing left to undo, redo now possible", !set.CanUndo && set.CanRedo, "");

        var (s3, _) = set.Redo();
        check("redo: restores the first edit", s3 == 0 && set.TryGet(0, 1, 0, out var g3) && Equals(g3, a), "");
        set.Redo();
        check("redo: restores the second edit", set.TryGet(0, 2, 0, out var g4) && Equals(g4, b), "");

        // A fresh action clears the redo stack.
        set.Undo();
        set.Apply(0, new List<(int, int, CellEdit?)> { (3, 0, a) });
        check("apply: a new action clears redo", !set.CanRedo, "");

        // after == null removes the edit and returns the cell to the file's value.
        set.Apply(0, new List<(int, int, CellEdit?)> { (1, 0, null) });
        check("apply: null removes the edit", !set.TryGet(0, 1, 0, out _), "");

        // A no-op (same value) is not recorded.
        var before = set.Apply(0, new List<(int, int, CellEdit?)> { (7, 7, a) });
        var again = set.Apply(0, new List<(int, int, CellEdit?)> { (7, 7, a) });
        check("apply: the same value is not a second change", before.Count == 1 && again.Count == 0, again.Count.ToString());

        set.Clear();
        check("clear: forgets edits and stacks", !set.Dirty && !set.CanUndo && !set.CanRedo, "");

        // Widths and freeze make the set dirty even with no cell edit.
        var w = new EditSet();
        w.SetColWidth(0, 2, 120);
        w.SetFreeze(0, 1, 0);
        check("dirty: a width and freeze count as changes", w.Dirty && w.ColWidths[(0, 2)] == 120 && w.Freeze[0] == (1, 0), "");
    }

    static void OnePasteOneRecord(Action<string, bool, string> check)
    {
        var set = new EditSet();
        var cells = new List<(int, int, CellEdit?)>();
        for (int r = 0; r < 3; r++)
            for (int c = 0; c < 2; c++)
                cells.Add((r, c, new CellEdit(EditKind.Number, Num: r * 10 + c)));

        var changed = set.Apply(4, cells);
        check("paste: 3 x 2 makes 6 edits", changed.Count == 6 && set.All.Count == 6, set.All.Count.ToString());
        check("paste: one undo record", set.CanUndo, "");

        var (sheet, undone) = set.Undo();
        check("paste: one undo reverses all 6 cells", sheet == 4 && undone.Count == 6 && set.All.Count == 0, set.All.Count.ToString());
        check("paste: one undo, then nothing left", !set.CanUndo, "");
    }
}
