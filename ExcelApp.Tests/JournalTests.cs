// Auto-save journal (EditJournal): snapshot, write, find, restore, forget, a changed file and a broken journal.
// Made-up data only, under D:\ExcelAppTest\syn-journal.
using ExcelApp;
using ExcelApp.Data;

static class JournalTests
{
    public static void Run(Action<string, bool, string> check)
    {
        string dir = @"D:\ExcelAppTest\syn-journal";
        if (Directory.Exists(dir)) Directory.Delete(dir, true);
        var store = StoreRoot.ForTest(dir);
        string file = Path.Combine(dir, "made-up.xlsx");
        File.WriteAllText(file, "made-up bytes, never parsed");

        var edits = new EditSet();
        edits.Apply(0, new List<(int, int, CellEdit?)> { (1, 2, new CellEdit(EditKind.Number, Num: 42.5)), (3, 0, new CellEdit(EditKind.Text, Text: "abc", Style: new StyleDelta(Bold: true))) });
        edits.Apply(1, new List<(int, int, CellEdit?)> { (0, 0, new CellEdit(EditKind.Formula, Formula: "=A1*2")), (9999, 0, new CellEdit(EditKind.Bool, Num: 1, Bool: true)) });
        edits.SetColWidth(0, 2, 120); edits.SetFreeze(1, 1, 0);

        var j = new EditJournal(store) { DelayMs = 5000 };
        j.Schedule(() => EditJournal.Take(file, edits));
        check("journal: nothing written before the quiet time", !store.Exists(EditJournal.KeyOf(file)), "");
        j.Flush();
        check("journal: written to the store", store.Exists(EditJournal.KeyOf(file)), "");

        var s = j.Find(file);
        check("journal: found", s != null && s.Cells.Count == 4 && s.Widths.Count == 1 && s.Freezes.Count == 1, $"{s?.Cells.Count}");
        check("journal: same file", s != null && EditJournal.SameFile(s), "");

        var back = new EditSet();
        int n = EditJournal.Restore(s!, back, sh => sh switch { 0 => (100, 10), 1 => (50, 5), _ => (-1, -1) });
        check("journal: restore skips cells outside the sheet", n == 3, $"{n}");
        check("journal: number back", back.TryGet(0, 1, 2, out var e1) && e1.Kind == EditKind.Number && e1.Num == 42.5, "");
        check("journal: text and style back", back.TryGet(0, 3, 0, out var e2) && e2.Text == "abc" && e2.Style?.Bold == true, "");
        check("journal: formula back", back.TryGet(1, 0, 0, out var e3) && e3.Kind == EditKind.Formula && e3.Formula == "=A1*2", "");
        check("journal: width and freeze back", back.ColWidths.Count == 1 && back.Freeze.TryGetValue(1, out var fz) && fz.Rows == 1, "");
        check("journal: restore is undoable", back.CanUndo && back.Undo().Sheet >= 0, "");

        File.AppendAllText(file, " changed");
        check("journal: changed file detected", !EditJournal.SameFile(j.Find(file)!), "");

        var empty = new EditSet();
        j.Schedule(() => EditJournal.Take(file, empty)); j.Flush();
        check("journal: no edits removes the journal", !store.Exists(EditJournal.KeyOf(file)), "");

        j.Schedule(() => EditJournal.Take(file, edits)); j.Flush();
        j.Forget(file);
        check("journal: forget removes it", !store.Exists(EditJournal.KeyOf(file)) && j.Find(file) == null, "");

        j.Schedule(() => EditJournal.Take(file, edits)); j.Forget(file); j.Flush();
        check("journal: forget cancels a pending write", !store.Exists(EditJournal.KeyOf(file)), "");

        store.WriteAllText(EditJournal.KeyOf(file), "{ not json");
        check("journal: broken journal is dropped", j.Find(file) == null && !store.Exists(EditJournal.KeyOf(file)), "");

        check("journal: key differs per file", EditJournal.KeyOf(file) != EditJournal.KeyOf(file + "x") && EditJournal.KeyOf(file) == EditJournal.KeyOf(file.ToUpperInvariant()), "");
        j.Dispose();
        Directory.Delete(dir, true);
    }
}
