# Plan 5: insert and delete rows and columns (the owner's next feature, 2026-11-14)

Order chosen by the owner: insert and delete rows and columns, then charts and pivot tables, then the dashboard designer.

## Why it is not a small change
The editor shows a read-only fast copy (`.xvc`) with an edit set on top, both addressed by **cache row and column**
(`Session.Real(st, view)`, then `Cells.Render(sheet, reader, row, col)`). Inserting a row means every position below it
moves, in the edits, the stored formats, the formulas and the save. Found by search on 2026-11-14:

| Place | What uses positions |
|---|---|
| `Session.cs` | 15 view-to-cache conversions, 14 cache reads (rows, cell detail, copy, fill, find, stats, edge) |
| `Cells.cs` | `Render`, `Raw`, `ValueEdit`, `CurrentFormat`, `SheetSource.ReadCore` (filter and sort) |
| `EditSet.cs` | edits keyed by (sheet, row, col), column widths, freeze, undo and redo records |
| `Formula/Calc.cs` | the formula grid reads cells by position; references must move as Excel moves them |
| `StyleStore.cs` | formats from the file, by cache position (36 uses) |
| `EditJournal.cs` | the unsaved-edits journal stores positions |
| `SaveBack.cs` | writes edits by address through Excel |
| `SessionAgent.cs` | agent reads and writes by address |

## The design
1. **One layout per sheet** (`SheetLayout`): `int[]? RowOf` maps a logical row to its cache row (-1 = a new blank row)
   and `int[]? ColOf` does the same for columns. Null means unchanged, so files that are never restructured pay nothing.
   1,000,000 rows cost 4 MB.
2. **Logical positions everywhere above the cache.** The edit set, the view map (filter and sort), formulas, the
   journal and the agent all use logical rows and columns. Exactly two places translate to cache positions: `Cells`
   (and its `SheetSource`) when reading the fast copy, and `StyleStore` when reading the file's formats.
3. **Insert n rows at r:** `RowOf` gets n entries of -1 at r. Edit keys with row >= r move down by n. Every edited
   formula has its references moved the way Excel moves them: references at or below r move, and a range that spans r
   grows. File formulas are read through the layout, so they point at the right cells. **Delete** removes the entries.
   Edits inside the deleted rows go away, references into them become `#REF!`, and ranges shrink.
4. **One undo step** per insert or delete: the record keeps the old layout slice and the edits it removed, so undo
   restores both exactly.
5. **Save:** `SaveBack` first replays the inserts and deletes in Excel, in order (`Rows(r).Resize(n).EntireRow.Insert`
   or `.Delete`), so Excel moves its own formulas, names and formats. Then it writes the edits at their logical
   addresses. After the save the fast copy is read again and the layout resets.
6. **Interface:** the ribbon's Insert and Delete buttons (now disabled) and a right-click menu on row and column
   headers. Ctrl++ and Ctrl+- as in Excel. The agent gets the same through `run_command` (`insertRows`,
   `deleteRows`, `insertCols`, `deleteCols`).

## Tests
- **Offline:** layout mapping, edit-key moves, formula reference moves for insert and delete (#REF!, ranges growing and
  shrinking), undo and redo, and journal round trips. These run on any machine.
- **Live on the laptop (`check.ps1`):** a made-up workbook. Insert and delete rows and columns, check the cells,
  formulas and formats, save through Excel, read the saved file back with Excel, and compare.
