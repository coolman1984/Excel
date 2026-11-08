using System.Globalization;
using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data.Query;

public sealed record TileValue(string TileId, double? Value, string Text, double? Previous, double? DeltaPct, DateTime AsOf, string? Error = null);
public sealed class Dashboards : IDisposable
{
    readonly DataServices services; readonly object gate = new(); readonly Dictionary<string, IReadOnlyList<TileValue>> cache = new();
    public Dashboards(DataServices services) { this.services = services; if (services.Ingestion != null) services.Ingestion.Loaded += Loaded; }
    static string PathFor(string id) { if (string.IsNullOrWhiteSpace(id) || id.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '-' or '_'))) throw new UserError("That dashboard identifier is invalid."); return "dashboards/" + id + ".json"; }
    public IReadOnlyList<Dashboard> List() => services.Config.List("dashboards", "*.json").Select(p => JsonSerializer.Deserialize<Dashboard>(services.Config.ReadAllText(p), Json.Options)!).ToArray();
    public Dashboard Get(string id) { string p = PathFor(id); return services.Config.Exists(p) ? JsonSerializer.Deserialize<Dashboard>(services.Config.ReadAllText(p), Json.Options)! : throw new UserError("That dashboard does not exist."); }
    public Dashboard Save(Dashboard dashboard) { services.Config.WriteAllText(PathFor(dashboard.Id), JsonSerializer.Serialize(dashboard, Json.Options)); lock (gate) cache.Remove(dashboard.Id); return dashboard; }
    public void Delete(string id) { services.Config.Delete(PathFor(id)); lock (gate) cache.Remove(id); }
    void Loaded(LoadRecord record) => _ = Task.Run(() => { foreach (var dashboard in List()) if (dashboard.Tiles.Any(t => Views.Depends(t.Spec, record.RecipeId))) try { Refresh(dashboard.Id); } catch (Exception) { } });
    static QuerySpec Clone(QuerySpec spec) => JsonSerializer.Deserialize<QuerySpec>(JsonSerializer.Serialize(spec, Json.Options), Json.Options)!;
    double? Scalar(QuerySpec spec, string measure) { var t = Engine.Run(spec, services.Tables, CancellationToken.None).Table; if (measure == "rows") return t.Rows; var c = t[measure]; return t.Rows == 0 || c == null || c.IsNull(0) || c.Type == ColType.Text ? null : c.Num(0); }
    QuerySpec Previous(Tile tile, DateTime now)
    {
        var spec = Clone(tile.Spec);
        if (tile.CompareTo == "previousVersion")
        {
            var info = services.Tables.List().FirstOrDefault(x => x.Kind == spec.Source.Kind && x.Id == spec.Source.Id) ?? throw new UserError("The previous version is unavailable.");
            spec.Source = spec.Source with { AsOf = info.LoadedAt.AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture), Version = null }; return spec;
        }
        if (tile.CompareTo == "previousPeriod")
        {
            var recipe = services.Ingestion.Recipes.FirstOrDefault(r => r.Id == spec.Source.Id) ?? throw new UserError("The report recipe is unavailable."); DateTime start, end;
            if (recipe.Granularity == Granularity.Month) { end = new DateTime(now.Year, now.Month, 1); start = end.AddMonths(-1); }
            else if (recipe.Granularity == Granularity.Day) { end = now.Date; start = end.AddDays(-1); }
            else { int offset = ((int)now.DayOfWeek - (int)recipe.Week.FirstDay + 7) % 7; end = now.Date.AddDays(-offset); start = end.AddDays(-7); }
            var predicates = new List<object>(); if (spec.Where is { } where) predicates.Add(where); predicates.Add(new { op = "ge", col = recipe.DateColumn, value = start.ToOADate() }); predicates.Add(new { op = "lt", col = recipe.DateColumn, value = end.ToOADate() }); spec.Where = JsonSerializer.SerializeToElement(new { op = "and", items = predicates }, Json.Options); return spec;
        }
        throw new UserError("That comparison is unavailable.");
    }
    public IReadOnlyList<TileValue> Values(string id) { lock (gate) if (cache.TryGetValue(id, out var result)) return result; return Refresh(id); }
    public IReadOnlyList<TileValue> Refresh(string id)
    {
        var now = DateTime.Now; var values = new List<TileValue>(); foreach (var tile in Get(id).Tiles)
        {
            try { double? value = Scalar(tile.Spec, tile.Measure), previous = tile.CompareTo == null ? null : Scalar(Previous(tile, now), tile.Measure); double? delta = previous == null || previous == 0 || value == null ? null : (value - previous) / Math.Abs(previous.Value) * 100; values.Add(new TileValue(tile.Id, value, value?.ToString("G15", CultureInfo.InvariantCulture) ?? "", previous, delta, now)); }
            catch (Exception) { values.Add(new TileValue(tile.Id, null, "", null, null, now, "This tile could not be calculated.")); }
        }
        lock (gate) cache[id] = values; return values;
    }
    public void Dispose() { if (services.Ingestion != null) services.Ingestion.Loaded -= Loaded; }
}
