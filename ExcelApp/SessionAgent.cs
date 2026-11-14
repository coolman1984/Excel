// Commands for AI agents (owner's decision, 2026-11-14): read and write the open workbook by Excel address and sheet name,
// as a person would say it ("Sheet1!B2:D9"), instead of the interface's 0-based view coordinates.
using System.Globalization;
using System.Text.Json;
using ExcelApp.Agent;
using ExcelViewer.Core;

namespace ExcelApp;

public sealed partial class Session
{
    public const int AgentMaxCells = 20_000;

    void RegisterAgent(Dictionary<string, Func<JsonElement, Task<object?>>> h)
    {
        h["agent.read"] = a => Task.FromResult<object?>(AgentRead(a));
        h["agent.write"] = a => Task.FromResult<object?>(AgentWrite(a));
    }

    int AgentSheet(CacheData c, JsonElement a)
    {
        if (!a.TryGetProperty("sheet", out var s) || s.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined) return current;
        if (s.ValueKind == JsonValueKind.Number) { int i = s.GetInt32(); if (i >= 0 && i < c.Sheets.Count) return i; }
        else if (s.ValueKind == JsonValueKind.String)
        {
            string name = s.GetString()!;
            int i = c.Sheets.FindIndex(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
            if (i >= 0) return i;
        }
        throw new UserError($"There is no sheet {s}. Sheets: {string.Join(", ", c.Sheets.Select(x => x.Name))}.");
    }

    object AgentRead(JsonElement a)
    {
        gate.Wait();
        var c = Need(); int sheet = AgentSheet(c, a); var d = c.Sheets[sheet];
        var (r1, c1, r2, c2) = A1.Parse(Arg.Str(a, "range") ?? throw new UserError("Say which range to read, such as A1:D20."), d.FirstRow + d.Rows - 1, d.FirstCol + d.Cols - 1);
        long count = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
        if (count > AgentMaxCells) throw new UserError($"That range has {count:N0} cells; read at most {AgentMaxCells:N0} at a time (for example A1:T1000). For questions over all rows, use the data store's SQL.");
        var values = new string[r2 - r1 + 1][]; string?[][]? formulas = null;
        lock (readerLock)
        {
            var rd = Reader(sheet);
            for (int r = r1; r <= r2; r++)
            {
                var row = values[r - r1] = new string[c2 - c1 + 1]; int real = r - d.FirstRow;
                for (int k = c1; k <= c2; k++)
                {
                    int col = k - d.FirstCol;
                    if (real < 0 || real >= d.Rows || col < 0 || col >= d.Cols) { row[k - c1] = ""; continue; }
                    row[k - c1] = cells!.Render(sheet, rd, real, col).Text;
                    string raw = cells.Raw(sheet, rd, real, col);
                    if (raw.StartsWith('='))
                    {
                        formulas ??= Enumerable.Range(0, r2 - r1 + 1).Select(_ => new string?[c2 - c1 + 1]).ToArray();
                        formulas[r - r1][k - c1] = raw;
                    }
                }
            }
        }
        return new
        {
            sheet = d.Name, range = A1.Format(r1, c1, r2, c2), values, formulas,
            used = d.Rows == 0 || d.Cols == 0 ? "" : A1.Format(d.FirstRow, d.FirstCol, d.FirstRow + d.Rows - 1, d.FirstCol + d.Cols - 1),
        };
    }

    object AgentWrite(JsonElement a)
    {
        var c = Need(); int sheet = AgentSheet(c, a); var st = St(sheet); var d = c.Sheets[sheet];
        var (r0, c0, _, _) = A1.Parse(Arg.Str(a, "start") ?? throw new UserError("Say where to start writing, such as B2."));
        if (!a.TryGetProperty("values", out var rows) || rows.ValueKind != JsonValueKind.Array) throw new UserError("Give the values as rows, such as [[\"Name\", 5], [\"Total\", \"=SUM(B1:B1)\"]].");
        var list = new List<(int, int, CellEdit?)>(); long count = 0;
        string used = d.Rows == 0 || d.Cols == 0 ? "(empty)" : A1.Format(d.FirstRow, d.FirstCol, d.FirstRow + d.Rows - 1, d.FirstCol + d.Cols - 1);
        int i = 0;
        foreach (var row in rows.EnumerateArray())
        {
            if (row.ValueKind != JsonValueKind.Array) throw new UserError("Each row of values must be a list, such as [\"a\", 1].");
            int k = 0;
            foreach (var v in row.EnumerateArray())
            {
                if (++count > AgentMaxCells) throw new UserError($"Write at most {AgentMaxCells:N0} cells at a time.");
                int excelRow = r0 + i, excelCol = c0 + k, real = excelRow - d.FirstRow, col = excelCol - d.FirstCol;
                if (real < 0 || real >= d.Rows || col < 0 || col >= d.Cols)
                    throw new UserError($"{A1.Format(excelRow, excelCol, excelRow, excelCol)} is outside the sheet's data ({used}); this editor writes inside it only.");
                CellEdit edit = v.ValueKind switch
                {
                    JsonValueKind.Null => new CellEdit(EditKind.Clear),
                    JsonValueKind.Number => KeepStyle(sheet, real, col, new CellEdit(EditKind.Number, Num: v.GetDouble())),
                    JsonValueKind.True or JsonValueKind.False => KeepStyle(sheet, real, col, new CellEdit(EditKind.Bool, Num: v.GetBoolean() ? 1 : 0, Bool: v.GetBoolean())),
                    JsonValueKind.String => Typed(sheet, real, col, v.GetString()!),
                    _ => throw new UserError("A value must be text, a number, true/false or null."),
                };
                list.Add((real, col, edit)); k++;
            }
            i++;
        }
        if (list.Count == 0) return new { written = 0 };
        Mutate(sheet, st, list);   // one action: one undo
        return new { written = list.Count, sheet = d.Name, dirty = Dirty, canUndo = Edits.CanUndo };
    }
}
