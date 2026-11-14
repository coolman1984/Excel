// Charts in the editor: chart.add / chart.update / chart.delete / chart.list. A chart is a range of a sheet plus a type;
// its points are read live from the cells (formulas included), so it follows every edit. Saving adds it to the file as an
// Excel chart (SaveBack). Undo, the journal and inserted or deleted rows treat charts like other edits (EditSet).
using System.Globalization;
using System.Text.Json;
using ExcelApp.Agent;

namespace ExcelApp;

public sealed partial class Session
{
    void RegisterCharts(Dictionary<string, Func<JsonElement, Task<object?>>> h)
    {
        h["chart.list"] = a => Task.FromResult<object?>(ChartList(a));
        h["chart.add"] = a => Task.FromResult<object?>(ChartAdd(a));
        h["chart.update"] = a => Task.FromResult<object?>(ChartUpdate(a));
        h["chart.delete"] = a => Task.FromResult<object?>(ChartDelete(a));
    }

    static string ChartType(string? t) => t != null && ChartData.Types.Contains(t) ? t : throw new UserError("Choose a chart type: column, bar, line, area, pie or scatter.");

    object ChartAdd(JsonElement a)
    {
        var c = Need(); int sheet = Arg.Int(a, "sheet"); var st = St(sheet); var d = c.Sheets[sheet];
        if (st.Map != null) throw new UserError("Clear the filter and the sorting on this sheet first, then make the chart.");
        string type = ChartType(Arg.Str(a, "type") ?? "column");
        (int R1, int C1, int R2, int C2) r;
        if (Arg.Str(a, "address") is string addr)   // the agent says "A1:C13"
        {
            var (x1, y1, x2, y2) = A1.Parse(addr, d.FirstRow + RowsOf(sheet) - 1, d.FirstCol + ColsOf(sheet) - 1);
            r = (x1 - d.FirstRow, y1 - d.FirstCol, x2 - d.FirstRow, y2 - d.FirstCol);
        }
        else r = a.TryGetProperty("range", out var rg) ? Arg.Rect(rg) : throw new UserError("Select the cells to chart first.");
        int rows = RowsOf(sheet), cols = ColsOf(sheet);
        r = (Math.Max(0, r.R1), Math.Max(0, r.C1), Math.Min(rows - 1, r.R2), Math.Min(cols - 1, r.C2));
        if (r.R1 > r.R2 || r.C1 > r.C2) throw new UserError("Those cells are outside the sheet.");
        if (r.R1 == r.R2 && r.C1 == r.C2) throw new UserError("Select the data to chart: a header row and the numbers below it, such as A1:C13.");
        int count; lock (editLock) count = Edits.Charts.Count(x => x.Sheet == sheet);
        var spec = new ChartSpec(Guid.NewGuid().ToString("N")[..10], sheet, type, Arg.Str(a, "title"), r.R1, r.C1, r.R2, r.C2,
            Arg.Num(a, "x", 120 + 28 * (count % 8)), Arg.Num(a, "y", 90 + 28 * (count % 8)), Arg.Num(a, "w", 520), Arg.Num(a, "h", 320));
        lock (editLock) Edits.SetChart(null, spec);
        Dirtied();
        return new { chart = ChartWire(spec), dirty = Dirty, canUndo = Edits.CanUndo };
    }

    object ChartUpdate(JsonElement a)
    {
        Need(); string id = Arg.Str(a, "id") ?? throw new UserError("Say which chart.");
        ChartSpec old; lock (editLock) old = Edits.Chart(id) ?? throw new UserError("That chart does not exist any more.");
        var spec = old with
        {
            Type = Arg.Str(a, "type") is string t ? ChartType(t) : old.Type,
            Title = a.TryGetProperty("title", out var tt) ? (tt.ValueKind == JsonValueKind.String ? tt.GetString() : null) : old.Title,
            X = Math.Max(0, Arg.Num(a, "x", old.X)), Y = Math.Max(0, Arg.Num(a, "y", old.Y)),
            W = Math.Clamp(Arg.Num(a, "w", old.W), 200, 2000), H = Math.Clamp(Arg.Num(a, "h", old.H), 150, 1400),
        };
        if (spec == old) return new { chart = ChartWire(spec), dirty = Dirty };
        lock (editLock) Edits.SetChart(old, spec);
        Dirtied();
        return new { chart = ChartWire(spec), dirty = Dirty, canUndo = Edits.CanUndo };
    }

    object ChartDelete(JsonElement a)
    {
        Need(); string id = Arg.Str(a, "id") ?? throw new UserError("Say which chart.");
        lock (editLock) { var old = Edits.Chart(id) ?? throw new UserError("That chart does not exist any more."); Edits.SetChart(old, null); }
        Dirtied();
        return new { deleted = id, dirty = Dirty, canUndo = Edits.CanUndo };
    }

    object ChartList(JsonElement a)
    {
        gate.Wait();
        Need(); int sheet = Arg.Int(a, "sheet", -1);
        List<ChartSpec> list; lock (editLock) list = Edits.Charts.Where(x => sheet < 0 || x.Sheet == sheet).OrderBy(x => x.Id).ToList();
        return new { charts = list.Select(ChartWire).ToList() };
    }

    /// <summary>A chart as the interface draws it: its card, its range as Excel writes it, and its points read from the cells now.</summary>
    object ChartWire(ChartSpec s)
    {
        var d = cache!.Sheets[s.Sheet];
        int rows = Math.Min(s.R2 - s.R1 + 1, ChartData.MaxPoints + 1), cols = Math.Min(s.C2 - s.C1 + 1, ChartData.MaxSeries + 1);
        ChartPoints p;
        lock (readerLock)
        {
            var rd = Reader(s.Sheet);
            p = ChartData.Read(rows, cols, (r, c) => { var v = cells!.Render(s.Sheet, rd, s.R1 + r, s.C1 + c); return (v.Kind, v.Num, v.Text); }, s.Type);
        }
        bool more = p.Truncated || s.R2 - s.R1 + 1 > rows || s.C2 - s.C1 + 1 > cols;
        return new
        {
            id = s.Id, sheet = s.Sheet, type = s.Type, title = s.Title, x = s.X, y = s.Y, w = s.W, h = s.H,
            range = A1.Format(d.FirstRow + s.R1, d.FirstCol + s.C1, d.FirstRow + s.R2, d.FirstCol + s.C2),
            categories = p.Categories, series = p.Series.Select(x => new { name = x.Name, values = x.Values }).ToList(), xValues = p.X, truncated = more,
        };
    }
}
