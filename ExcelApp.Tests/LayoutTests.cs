// Inserted and deleted rows and columns as the editor shows them, over a small made-up fast copy (no Excel needed):
// cells, formulas, the formula bar, fill values, filtering and find all read through the layout.
using System.Text;
using ExcelApp;
using ExcelApp.Data;
using ExcelViewer.Core;

public static class LayoutTests
{
    /// <summary>A made-up fast copy: Sheet1, 6 rows x 3 columns: A = 1..6, B = "b1".."b6", C = 10..60. Excel rows start at 1.</summary>
    static string MakeCache(StoreRoot root)
    {
        string path = root.PathOf("syn-layout.xvc");
        var cache = new CacheData(); cache.Formats.Add(new FormatEntry { Code = "General", Status = FormatStatus.MatchesExcel });
        const int rows = 6, cols = 3;
        var d = new SheetInfo { Name = "Sheet1", FirstRow = 1, FirstCol = 1, Rows = rows, Cols = cols, BlockRows = 4, ColWidths = new float[cols], RunStart = new int[cols][], RunFmt = new int[cols][] };
        for (int k = 0; k < cols; k++) { d.ColWidths[k] = 10; d.RunStart[k] = new[] { 0 }; d.RunFmt[k] = new[] { 0 }; }
        cache.Sheets.Add(d);
        using (var fs = new FileStream(path, FileMode.Create)) using (var w = new BinaryWriter(fs, Encoding.UTF8))
        {
            w.Write(Encoding.ASCII.GetBytes(CacheData.Magic)); w.Write(0L); w.Write(0L);
            for (int r = 0; r < rows; r += d.BlockRows)
            {
                int n = Math.Min(d.BlockRows, rows - r); var off = new long[cols + 1];
                for (int k = 0; k < cols; k++)
                {
                    var kinds = new byte[n]; double[]? nums = k == 1 ? null : new double[n]; string[]? text = k == 1 ? new string[n] : null;
                    for (int i = 0; i < n; i++) { if (k == 1) { kinds[i] = CellKind.Text; text![i] = "b" + (r + i + 1); } else { kinds[i] = CellKind.Number; nums![i] = k == 0 ? r + i + 1 : (r + i + 1) * 10; } }
                    var cc = ColumnChunk.Build(n, kinds, nums, text); w.Flush(); off[k] = fs.Position; cc.Write(w);
                }
                w.Flush(); off[cols] = fs.Position; d.Chunks.Add(off);
            }
            w.Flush(); long at = fs.Position; cache.WriteDirectory(w); w.Write(at);
        }
        return path;
    }

    public static void Run(Action<string, bool, string> check)
    {
        var root = StoreRoot.ForTest(@"D:\ExcelAppTest\syn-layout-" + Guid.NewGuid().ToString("N")[..8]);
        var cache = CacheData.Load(MakeCache(root));
        var edits = new EditSet(); var gate = new object();
        edits.ShiftFormula = (fs, f, op) => ExcelApp.Formula.FormulaShift.Structural(f, fs == op.Sheet, "Sheet1", op.Rows, op.At + 1, op.Count, op.Insert);
        var cells = new Cells(cache, edits, gate, () => null);
        using var rd = new SheetReader(cache, cache.Sheets[0]);
        string T(int r, int c) => cells.Render(0, rd, r, c).Text;

        edits.Apply(0, new (int, int, CellEdit?)[] { (6 - 1, 2, new CellEdit(EditKind.Formula, Formula: "=SUM(C1:C5)")) });   // C6 = SUM(C1:C5) = 150
        check("layout: before any insert", T(0, 0) == "1" && T(5, 2) == "150" && cells.RowCount(0) == 6, T(5, 2));

        // insert 2 rows before Excel row 3 (logical 2)
        edits.Structure(new StructOp(0, true, 2, 2, true), 6, 3);
        check("layout: rows above stay", T(1, 0) == "2" && T(1, 1) == "b2", T(1, 0));
        check("layout: inserted rows are empty", T(2, 0) == "" && T(3, 1) == "" && cells.Render(0, rd, 2, 0).Kind == CellKind.Empty && cells.Raw(0, rd, 3, 2) == "", T(2, 0));
        check("layout: rows below show the file's next rows", T(4, 0) == "3" && T(4, 1) == "b3" && T(6, 2) == "50" && cells.RowCount(0) == 8, T(4, 0));
        check("layout: the formula moved and still adds the same cells", cells.TryEdit(0, 7, 2, out var f) && f.Formula == "=SUM(C1:C7)" && T(7, 2) == "150", T(7, 2));
        edits.Apply(0, new (int, int, CellEdit?)[] { (2, 2, new CellEdit(EditKind.Number, Num: 1000)) });
        check("layout: typing in an inserted row counts in the formula", T(7, 2) == "1150" && T(2, 2) == "1000", T(7, 2));
        check("layout: fill values and the file's own values read through the layout", cells.ValueEdit(0, rd, 4, 1)!.Text == "b3" && cells.ValueEdit(0, rd, 3, 1)!.Kind == EditKind.Clear && cells.FileValue(0, rd, 5, 0).Num == 4 && cells.FileValue(0, rd, 3, 0).Kind == CellKind.Empty, "");
        check("layout: a file position maps to where it now shows", cells.Logical(0, 2, 1) == (4, 1) && cells.Logical(0, 0, 0) == (0, 0), cells.Logical(0, 2, 1).ToString());

        // filtering reads through the layout too
        using var src = new SheetSource(cache, 0, cells);
        var fe = new FilterEngine(src) { HeaderRows = 0 };
        fe.SetCondition(1, new TextCondition("begins", "b"));
        var view = fe.Apply(CancellationToken.None);
        check("layout: a filter sees the logical rows", src.Rows == 8 && view != null && view.SequenceEqual(new[] { 0, 1, 4, 5, 6, 7 }), view == null ? "null" : string.Join(",", view));

        // delete the column B (logical 1): C moves left, B's text is gone
        edits.Structure(new StructOp(0, false, 1, 1, false), 6, 3);
        check("layout: a deleted column's cells are gone, the next column moves left", T(4, 1) == "30" && T(2, 1) == "1000" && cells.ColCount(0) == 2 && T(7, 1) == "1150", T(4, 1) + " " + T(7, 1));
        check("layout: the formula follows its column", cells.TryEdit(0, 7, 1, out var f2) && f2.Formula == "=SUM(B1:B7)", cells.TryEdit(0, 7, 1, out var f3) ? f3.Formula! : "none");
        check("layout: a deleted column's file cells map nowhere", cells.Logical(0, 0, 1) == (-1, -1) && cells.Logical(0, 0, 2) == (0, 1), "");
        var e2 = new FilterEngine(src) { HeaderRows = 0 }; e2.SetCondition(1, new NumberCondition("gt", 100));
        var v2 = e2.Apply(CancellationToken.None);
        check("layout: filters see the moved column", v2 != null && v2.SequenceEqual(new[] { 2, 7 }), v2 == null ? "null" : string.Join(",", v2));

        // undo both: the sheet is the file again
        edits.Undo(); edits.Undo(); edits.Undo();
        check("layout: undo returns to the file's layout", cells.RowCount(0) == 6 && cells.ColCount(0) == 3 && T(2, 0) == "3" && T(2, 1) == "b3" && T(5, 2) == "150" && !cells.HasLayout(0), T(2, 0));
    }
}
