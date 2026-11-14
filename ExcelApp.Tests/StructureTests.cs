// Inserting and deleting rows and columns, without Excel: references move as Excel moves them, edits move, undo is exact.
using ExcelApp;
using ExcelApp.Formula;

public static class StructureTests
{
    public static void Run(Action<string, bool, string> check)
    {
        // ---- formulas: insert 2 rows at row 3 of Sheet1 (the formula sits on Sheet1)
        string IR(string f) => FormulaShift.Structural(f, true, "Sheet1", true, 3, 2, true);
        string DR(string f) => FormulaShift.Structural(f, true, "Sheet1", true, 3, 2, false);   // delete rows 3:4
        string IC(string f) => FormulaShift.Structural(f, true, "Sheet1", false, 2, 1, true);   // insert column B
        string DC(string f) => FormulaShift.Structural(f, true, "Sheet1", false, 2, 1, false);  // delete column B
        void Eq(string name, string got, string want) => check("structure: " + name, got == want, got + " (want " + want + ")");
        Eq("insert: a cell above stays", IR("=A2*2"), "=A2*2");
        Eq("insert: a cell at or below moves", IR("=A3+$B$10"), "=A5+$B$12");
        Eq("insert: a range spanning the rows grows", IR("=SUM(A1:A5)"), "=SUM(A1:A7)");
        Eq("insert: a range below moves", IR("=SUM(B3:C4)"), "=SUM(B5:C6)");
        Eq("insert: whole columns stay, whole rows move", IR("=SUM(A:B)+SUM(3:4)+SUM(1:2)"), "=SUM(A:B)+SUM(5:6)+SUM(1:2)");
        Eq("insert: text and function names are left alone", IR("=IF(A4=\"A4\",LOG10(A4),0)"), "=IF(A6=\"A4\",LOG10(A6),0)");
        Eq("insert: other sheets' references stay", IR("=Data!A5+A5"), "=Data!A5+A7");
        Eq("insert: references to the sheet by name move", FormulaShift.Structural("=Sheet1!A5+'Sheet1'!B6+A5", false, "Sheet1", true, 3, 2, true), "=Sheet1!A7+'Sheet1'!B8+A5");
        Eq("insert: quoted sheet names with spaces", FormulaShift.Structural("='My Data'!C9", false, "My Data", true, 3, 1, true), "='My Data'!C10");
        Eq("delete: a cell in the deleted rows is #REF!", DR("=A3+A5"), "=#REF!+A3");
        Eq("delete: a range shrinks", DR("=SUM(A1:A10)"), "=SUM(A1:A8)");
        Eq("delete: a range wholly deleted is #REF!", DR("=SUM(A3:B4)"), "=SUM(#REF!)");
        Eq("delete: a range starting inside starts at the cut", DR("=SUM(A4:A9)"), "=SUM(A3:A7)");
        Eq("delete: whole rows", DR("=SUM(2:6)+SUM(3:4)"), "=SUM(2:4)+SUM(#REF!)");
        Eq("delete: a sheet-named reference keeps its name", FormulaShift.Structural("=Sheet1!A3", false, "Sheet1", true, 3, 2, false), "=Sheet1!#REF!");
        Eq("insert column: references from B move", IC("=A1+B1+$C$2+SUM(A1:C3)"), "=A1+C1+$D$2+SUM(A1:D3)");
        Eq("insert column: whole columns", IC("=SUM(B:C)+SUM(A:A)"), "=SUM(C:D)+SUM(A:A)");
        Eq("delete column: B is #REF!, C moves left", DC("=B1+C1+SUM(A1:C1)"), "=#REF!+B1+SUM(A1:B1)");
        Eq("numbers are not references", IR("=A5+3+1.5"), "=A7+3+1.5");

        // ---- the edit set: layout, moves, undo and redo
        var e = new EditSet { ShiftFormula = (sheet, f, op) => FormulaShift.Structural(f, sheet == op.Sheet, op.Sheet == 0 ? "Sheet1" : "Data", op.Rows, op.At + 1, op.Count, op.Insert) };
        e.Apply(0, new (int, int, CellEdit?)[] { (1, 0, new CellEdit(EditKind.Number, Num: 5)), (4, 0, new CellEdit(EditKind.Text, Text: "x")), (6, 1, new CellEdit(EditKind.Formula, Formula: "=SUM(A1:A6)")) });
        e.SetColWidth(0, 1, 90);
        e.Structure(new StructOp(0, true, 2, 3, true), 10, 4);   // insert 3 rows before logical row 2 (Excel row 3)
        check("edit set: rows grow and new rows are blank", e.RowCount(0, 10) == 13 && e.PhysRow(0, 1) == 1 && e.PhysRow(0, 2) == -1 && e.PhysRow(0, 4) == -1 && e.PhysRow(0, 5) == 2 && e.PhysRow(0, 12) == 9, "");
        check("edit set: the inverse map", e.LogicalRow(0, 2) == 5 && e.LogicalRow(0, 0) == 0, e.LogicalRow(0, 2).ToString());
        check("edit set: edits above stay, below move", e.TryGet(0, 1, 0, out var a) && a.Num == 5 && e.TryGet(0, 7, 0, out var b) && b.Text == "x" && !e.TryGet(0, 4, 0, out _), "");
        check("edit set: formulas move their references", e.TryGet(0, 9, 1, out var f) && f.Formula == "=SUM(A1:A9)", f?.Formula ?? "none");
        check("edit set: an insert is unsaved work", e.Dirty && e.Ops.Count == 1 && e.HasLayout(0) && !e.HasLayout(1), "");
        e.Apply(0, new (int, int, CellEdit?)[] { (3, 0, new CellEdit(EditKind.Text, Text: "new")) });
        e.Undo();                                                   // the typing
        var (us, _) = e.Undo();                                     // the insert
        check("edit set: undo puts back exactly the old state", us == 0 && e.LastWasStructural && e.RowCount(0, 10) == 10 && !e.HasLayout(0) && e.TryGet(0, 4, 0, out var b2) && b2.Text == "x" && e.TryGet(0, 6, 1, out var f2) && f2.Formula == "=SUM(A1:A6)" && e.Ops.Count == 0, "");
        e.Redo(); e.Redo();
        check("edit set: redo replays the insert and the typing", e.RowCount(0, 10) == 13 && e.TryGet(0, 3, 0, out var n) && n.Text == "new" && e.TryGet(0, 9, 1, out var f3) && f3.Formula == "=SUM(A1:A9)", "");
        e.Structure(new StructOp(0, true, 7, 1, false), 10, 4);    // delete the row holding "x"
        check("edit set: delete drops the row's edits and moves the rest up", !e.Edits().Any(x => x.Value.Text == "x") && e.TryGet(0, 8, 1, out var f4) && f4.Formula == "=SUM(A1:A8)" && e.RowCount(0, 10) == 12, e.TryGet(0, 8, 1, out var f5) ? f5.Formula! : "none");
        e.Structure(new StructOp(0, false, 0, 1, true), 10, 4);    // insert a column at A
        check("edit set: an inserted column moves edits and widths right", e.TryGet(0, 1, 1, out var c1) && c1.Num == 5 && e.ColWidths.ContainsKey((0, 2)) && !e.ColWidths.ContainsKey((0, 1)) && e.ColCount(0, 4) == 5 && e.PhysCol(0, 0) == -1 && e.PhysCol(0, 1) == 0, "");
        check("edit set: the formula moved with its cell and its references", e.TryGet(0, 8, 2, out var f6) && f6.Formula == "=SUM(B1:B8)", e.TryGet(0, 8, 2, out var f7) ? f7.Formula! : "none");
        check("edit set: the ops are kept in order for the save", e.Ops.Count == 3 && e.Ops[0].Insert && !e.Ops[1].Insert && !e.Ops[2].Rows, "");
        string bad; try { e.Structure(new StructOp(0, true, 11, 5, false), 10, 4); bad = "accepted"; } catch (ArgumentOutOfRangeException) { bad = "refused"; }
        check("edit set: deleting past the end is refused", bad == "refused", bad);
        // the unsaved-edits journal keeps the inserts and deletes and replays them before the cells
        var snap = EditJournal.Take(Path.Combine(Path.GetTempPath(), "syn-not-there.xlsx"), e);
        var json = System.Text.Json.JsonSerializer.Serialize(snap);
        var back = System.Text.Json.JsonSerializer.Deserialize<EditJournal.Snapshot>(json)!;
        var r = new EditSet { ShiftFormula = e.ShiftFormula };
        EditJournal.Restore(back, r, sh => sh == 0 ? (10, 4) : (-1, -1));
        check("journal: inserts and deletes come back, then the cells at their places", r.Ops.Count == 3 && r.RowCount(0, 10) == 12 && r.ColCount(0, 4) == 5 && r.TryGet(0, 8, 2, out var jf) && jf.Formula == "=SUM(B1:B8)" && r.TryGet(0, 1, 1, out var jn) && jn.Num == 5, r.Ops.Count.ToString());
        var old = System.Text.Json.JsonSerializer.Deserialize<EditJournal.Snapshot>("{\"File\":\"x\",\"Length\":1,\"WriteUtc\":\"2026-01-01T00:00:00Z\",\"SavedAt\":\"2026-01-01T00:00:00\",\"Cells\":[],\"Widths\":[],\"Freezes\":[]}")!;
        check("journal: an older journal without inserts still reads", old.Ops == null, "");
        e.Clear();
        check("edit set: a save clears the layout", !e.HasLayout(0) && e.Ops.Count == 0 && e.RowCount(0, 10) == 10 && !e.Dirty, "");
    }

    static IEnumerable<KeyValuePair<(int Sheet, int Row, int Col), CellEdit>> Edits(this EditSet e) => e.All;
}
