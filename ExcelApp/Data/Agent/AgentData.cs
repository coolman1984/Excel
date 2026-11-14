// Read-only access to the store for an AI agent (owner's decision, 2026-11-14: AI may read all the data).
// Nothing here writes: it lists tables, describes them and answers SQL through the Data mode's engine.
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelApp.Data.Query;

namespace ExcelApp.Data.Agent;

public sealed record AgentColumn(string Name, string Type);
public sealed record AgentResult(IReadOnlyList<AgentColumn> Columns, IReadOnlyList<object?[]> Rows, int RowCount, bool Truncated, long Ms);

public sealed class AgentData
{
    readonly ITableProvider tables;
    public AgentData(ITableProvider tables) => this.tables = tables;

    /// <summary>Compact camelCase JSON for agents (dates as ISO text, nulls kept).</summary>
    // Relaxed escaping: Arabic text and apostrophes stay readable (this JSON never goes into HTML).
    public static readonly JsonSerializerOptions JsonOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = false, DefaultIgnoreCondition = JsonIgnoreCondition.Never, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
    public const int DefaultRows = 1000, MaxRows = 100_000;

    static string TypeName(ColType t) => t.ToString().ToLowerInvariant();

    public object Tables() => new
    {
        tables = tables.List().Select(t => new
        {
            name = t.Id, kind = t.Kind, rows = t.Rows,
            loaded = t.LoadedAt == default ? null : t.LoadedAt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
            columns = t.Columns.Where(c => !c.Name.StartsWith('_')).Select(c => new { name = c.Name, type = TypeName(c.Type) }).ToArray()
        }).ToArray()
    };

    public object Describe(string table, int sampleRows = 5)
    {
        if (string.IsNullOrWhiteSpace(table)) throw new UserError("Name the table to describe.");
        var info = Sql.Plan("SELECT * FROM " + Quote(table) + " LIMIT 1", tables.List()).Spec.Source;
        var t = tables.List().First(x => x.Id == info.Id && x.Kind == info.Kind);
        var sample = Query("SELECT * FROM " + Quote(table), Math.Clamp(sampleRows, 0, 50));
        return new { name = t.Id, kind = t.Kind, rows = t.Rows, columns = t.Columns.Where(c => !c.Name.StartsWith('_')).Select(c => new { name = c.Name, type = TypeName(c.Type) }).ToArray(), sample = new { columns = sample.Columns.Select(c => c.Name), rows = sample.Rows } };
    }
    static string Quote(string name) => "\"" + name.Replace("\"", "") + "\"";

    /// <summary>Runs one SELECT. At most <paramref name="maxRows"/> rows come back; RowCount says how many matched.</summary>
    public AgentResult Query(string sql, int maxRows = DefaultRows, CancellationToken token = default)
    {
        var sw = Stopwatch.StartNew();
        maxRows = Math.Clamp(maxRows, 0, MaxRows);
        var plan = Sql.Plan(sql, tables.List());
        var spec = plan.Spec; int wanted = plan.Limit ?? int.MaxValue;
        spec.Limit = Math.Min(wanted, maxRows);
        QueryResult result;
        try { result = Engine.Run(spec, tables, token); }
        catch (UserError) { throw; }
        catch (Exception e) when (e is ArgumentException or FormatException or KeyNotFoundException or InvalidOperationException)
        { throw new UserError("The query could not run: " + e.Message); }
        var table = result.Table;
        var shown = (plan.Output ?? table.Columns.Select(c => (c.Name, c.Name)).ToList()).Where(o => table[o.Internal] != null && !(o.Internal.StartsWith('_') && o.Internal == o.Shown)).ToList();
        var cols = shown.Select(o => table[o.Internal]!).ToArray();
        var rows = new List<object?[]>(table.Rows);
        for (int r = 0; r < table.Rows; r++)
        {
            var row = new object?[cols.Length];
            for (int k = 0; k < cols.Length; k++) row[k] = Value(cols[k], r);
            rows.Add(row);
        }
        int count = Math.Min(result.Total, wanted);
        return new AgentResult(shown.Select((o, k) => new AgentColumn(o.Shown, TypeName(cols[k].Type))).ToArray(), rows, count, count > rows.Count, sw.ElapsedMilliseconds);
    }

    static object? Value(Column c, int r)
    {
        if (c.IsNull(r)) return null;
        switch (c.Type)
        {
            case ColType.Text: return c.Text(r);
            case ColType.Bool: return c.Num(r) != 0;
            case ColType.Date:
                double v = c.Num(r);
                if (v < -657434 || v > 2958465) return v;
                var d = DateTime.FromOADate(v);
                return d.TimeOfDay == TimeSpan.Zero ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : d.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
            default:
                double n = c.Num(r); return double.IsFinite(n) ? Cells.Excel15(n) : null;   // 15 significant digits, as Excel shows
        }
    }
}
