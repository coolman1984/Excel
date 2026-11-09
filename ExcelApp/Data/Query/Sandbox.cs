using System.Globalization;
using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data.Query;

public sealed class SandboxEdit { public int Row { get; set; } public string Col { get; set; } = ""; public JsonElement Value { get; set; } }
public sealed class SandboxData { public string Id { get; set; } = ""; public string Name { get; set; } = ""; public TableRef Base { get; set; } = new("table", ""); public DateTime CreatedAt { get; set; } public List<SandboxEdit> Changes { get; set; } = new(); public int Edits => Changes.Count; }
public sealed record SandboxDiff(string ResultId, int Rows, IReadOnlyList<ColumnInfo> Columns, double Ms, int Added, int Removed, int Changed, IReadOnlyDictionary<string, int> ChangedColumns);
public sealed class Sandboxes
{
    readonly DataServices services; readonly ITableProvider source; readonly object gate = new();
    public Sandboxes(DataServices services, ITableProvider source) { this.services = services; this.source = source; }
    static string PathFor(string id) { if (string.IsNullOrWhiteSpace(id) || id.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '-' or '_'))) throw new UserError("That sandbox identifier is invalid."); return "sandboxes/" + id + ".json"; }
    SandboxData Load(string id)
    {
        string path = PathFor(id);
        if (!services.Store.Exists(path)) throw new UserError("That sandbox does not exist.");
        var data = JsonSerializer.Deserialize<SandboxData>(services.Store.ReadAllText(path), Json.Options)!;
        if (data.Changes.Count > 0 && services.Store.Exists("sandboxes/" + id + ".xvt"))
        {
            var edited = TableFile.Open(services.Store, "sandboxes/" + id + ".xvt").Load();
            foreach (var e in data.Changes)
            {
                var c = edited[e.Col] ?? throw new UserError("The sandbox edits are damaged.");
                e.Value = JsonSerializer.SerializeToElement(c.IsNull(e.Row) ? null : c.Type switch {
                    ColType.Text => (object)c.Text(e.Row), ColType.Bool => c.Num(e.Row) != 0, _ => c.Num(e.Row) });
            }
        }
        return data;
    }
    void Save(SandboxData data)
    {
        if (data.Changes.Count > 0)
        {
            var edited = Apply(data);
            var names = data.Changes.Select(e => e.Col).ToHashSet(StringComparer.OrdinalIgnoreCase);
            TableFile.Write(services.Store, "sandboxes/" + data.Id + ".xvt", new Table(data.Name, edited.Columns.Where(c => names.Contains(c.Name)).ToArray()), new Dictionary<string,string>());
        }
        services.Store.WriteAllText(PathFor(data.Id), JsonSerializer.Serialize(new { data.Id, data.Name, data.Base, data.CreatedAt, changes=data.Changes.Select(e => new { e.Row, e.Col }) }, Json.Options));
    }
    public SandboxData Create(TableRef reference, string name) { var now = DateTime.Now; if (reference.Kind == "report" && reference.AsOf == null && reference.Version == null) reference = reference with { AsOf = now.ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture) }; source.Get(reference); var data = new SandboxData { Id = Guid.NewGuid().ToString("N"), Name = name, Base = reference, CreatedAt = now.ToUniversalTime() }; lock (gate) Save(data); return data; }
    public IReadOnlyList<SandboxData> List() { lock (gate) return services.Store.List("sandboxes", "*.json").Select(p => JsonSerializer.Deserialize<SandboxData>(services.Store.ReadAllText(p), Json.Options)!).ToArray(); }
    public void Discard(string id) { lock (gate) { services.Store.Delete(PathFor(id)); services.Store.Delete("sandboxes/" + id + ".xvt"); } }
    public int Edit(string id, IReadOnlyList<SandboxEdit> edits)
    {
        lock (gate)
        {
            var data = Load(id); var table = source.Get(data.Base); var map = data.Changes.ToDictionary(e => (e.Row, e.Col.ToUpperInvariant()));
            foreach (var edit in edits)
            {
                var col = table[edit.Col] ?? throw new UserError("That column does not exist."); if (edit.Row < 0 || edit.Row >= table.Rows || col.Name.StartsWith('_')) throw new UserError("That cell cannot be edited.");
                Validate(col, edit.Value); map[(edit.Row, col.Name.ToUpperInvariant())] = new SandboxEdit { Row = edit.Row, Col = col.Name, Value = edit.Value.Clone() };
            }
            data.Changes = map.Values.ToList(); Save(data); return data.Edits;
        }
    }
    static void Validate(Column column, JsonElement value)
    {
        if (value.ValueKind is JsonValueKind.Null) return;
        bool ok = column.Type switch { ColType.Text => value.ValueKind == JsonValueKind.String, ColType.Bool => value.ValueKind is JsonValueKind.True or JsonValueKind.False, ColType.Number => value.ValueKind == JsonValueKind.Number, ColType.Date => value.ValueKind == JsonValueKind.Number || (value.ValueKind == JsonValueKind.String && DateTime.TryParse(value.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out _)), _ => false };
        if (!ok) throw new UserError("The value does not match the column type.");
    }
    public Table Get(string id) { lock (gate) return Apply(Load(id)); }
    Table Apply(SandboxData data)
    {
        var table = source.Get(data.Base); var changes = data.Changes.GroupBy(e => e.Col, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.ToArray(), StringComparer.OrdinalIgnoreCase);
        var columns = new List<Column>(); foreach (var col in table.Columns)
        {
            if (!changes.TryGetValue(col.Name, out var edits)) { columns.Add(col); continue; }
            if (col.Type == ColType.Text) { var values = Enumerable.Range(0, table.Rows).Select(r => col.IsNull(r) ? null : col.Text(r)).ToArray(); foreach (var e in edits) values[e.Row] = e.Value.ValueKind == JsonValueKind.Null ? null : e.Value.GetString(); columns.Add(TextColumn.From(col.Name, values)); }
            else if (col.Type == ColType.Bool) { var values = Enumerable.Range(0, table.Rows).Select(r => col.IsNull(r) ? (sbyte)-1 : (sbyte)col.Num(r)).ToArray(); foreach (var e in edits) values[e.Row] = e.Value.ValueKind == JsonValueKind.Null ? (sbyte)-1 : e.Value.GetBoolean() ? (sbyte)1 : (sbyte)0; columns.Add(new BoolColumn(col.Name, values)); }
            else { var values = Enumerable.Range(0, table.Rows).Select(col.Num).ToArray(); var nulls = Enumerable.Range(0, table.Rows).Select(col.IsNull).ToArray(); foreach (var e in edits) { nulls[e.Row] = e.Value.ValueKind == JsonValueKind.Null; if (!nulls[e.Row]) values[e.Row] = e.Value.ValueKind == JsonValueKind.String ? DateTime.Parse(e.Value.GetString()!, CultureInfo.InvariantCulture).ToOADate() : e.Value.GetDouble(); } columns.Add(new NumberColumn(col.Name, values, nulls, col.Type)); }
        }
        return new Table(data.Name, columns);
    }
    public int Keep(string id, string reportType, string period, string note) => source.SaveAsVersion(reportType, period, Get(id), note);
    public SandboxDiff Compare(string id)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew(); SandboxData data; lock (gate) data = Load(id); var old = source.Get(data.Base); var edited = Get(id); var rows = new List<int>(); var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var rowEdits in data.Changes.GroupBy(x => x.Row).OrderBy(g => g.Key))
        {
            int row = rowEdits.Key; bool changed = false; foreach (var e in rowEdits) { var a = old[e.Col]!; var b = edited[e.Col]!; if (a.IsNull(row) != b.IsNull(row) || (!a.IsNull(row) && a.Text(row) != b.Text(row))) { changed = true; counts.TryGetValue(e.Col, out int n); counts[e.Col] = n + 1; } } if (changed) rows.Add(row);
        }
        var indexes = rows.ToArray(); var columns = edited.Columns.Where(c => !c.Name.StartsWith('_')).Select(c => Engine.Gather(c, indexes)).ToList();
        foreach (var name in counts.Keys) { var col = old[name]!; var values = indexes.Select(r => col.IsNull(r) ? null : col.Text(r)).ToArray(); columns.Add(TextColumn.From(name + " (before)", values)); }
        columns.Add(TextColumn.From("_status", Enumerable.Repeat<string?>("changed", indexes.Length).ToArray())); var result = new Table(data.Name, columns); string resultId = services.Results.Add(result, new ResultMeta("diff", null, false, data.Base)); watch.Stop();
        return new SandboxDiff(resultId, result.Rows, result.Columns.Select(c => new ColumnInfo(c.Name, c.Type)).ToArray(), watch.Elapsed.TotalMilliseconds, 0, 0, rows.Count, counts);
    }
}
public sealed class SandboxProvider : ITableProvider
{
    readonly ITableProvider inner; readonly Sandboxes sandboxes; public SandboxProvider(ITableProvider inner, Sandboxes sandboxes) { this.inner = inner; this.sandboxes = sandboxes; }
    public Table Get(TableRef r) => r.Kind == "sandbox" ? sandboxes.Get(r.Id) : inner.Get(r);
    public IReadOnlyList<TableInfo> List() => inner.List(); public SourceInfo? SourceOf(int id) => inner.SourceOf(id);
    public int SaveAsVersion(string reportTypeId, string period, Table data, string note) => inner.SaveAsVersion(reportTypeId, period, data, note);
}
