using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data.Query;

public sealed class Tile { public string Id { get; set; } = ""; public string Title { get; set; } = ""; public QuerySpec Spec { get; set; } = new(); public string Measure { get; set; } = ""; public string Format { get; set; } = ""; public string? CompareTo { get; set; } }
public sealed class Dashboard { public string Id { get; set; } = ""; public string Name { get; set; } = ""; public List<Tile> Tiles { get; set; } = new(); }
public sealed record LibrarySelection(string Kind, string Id);
public sealed record LibraryLiteral(string Kind, string Id, string Field, object? Value);
public sealed class LibraryItem { public string Kind { get; set; } = ""; public string Id { get; set; } = ""; public JsonElement Data { get; set; } }
public sealed class LibraryFile { public string Format { get; set; } = "exlib1"; public List<LibraryItem> Items { get; set; } = new(); }
public sealed class Library
{
    readonly DataServices services; readonly Views views; readonly Alerts alerts; readonly ReportRecipes recipes;
    public Library(DataServices s, Views v, Alerts a, ReportRecipes r) { services = s; views = v; alerts = a; recipes = r; }
    static string SafeId(string id) { if (string.IsNullOrWhiteSpace(id) || id.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '-' or '_'))) throw new UserError("That identifier is invalid."); return id; }
    Dashboard GetDashboard(string id) { string path = "dashboards/" + SafeId(id) + ".json"; return services.Config.Exists(path) ? JsonSerializer.Deserialize<Dashboard>(services.Config.ReadAllText(path), Json.Options)! : throw new UserError("That dashboard does not exist."); }
    public LibraryFile Build(IReadOnlyList<LibrarySelection> selections)
    {
        var result = new LibraryFile(); foreach (var item in selections)
        {
            object data = item.Kind switch { "view" => views.Get(item.Id), "link" => views.GetLink(item.Id), "alert" => alerts.List().First(x => x.Id == item.Id), "reportRecipe" => recipes.Get(item.Id), "dashboard" => GetDashboard(item.Id), "ingestRecipe" => CloneIngest(item.Id), _ => throw new UserError("That library item kind is unknown.") };
            result.Items.Add(new LibraryItem { Kind = item.Kind, Id = item.Id, Data = JsonSerializer.SerializeToElement(data, data.GetType(), Json.Options) });
        }
        return result;
    }
    IngestRecipe CloneIngest(string id)
    {
        var original = services.Ingestion.Recipes.FirstOrDefault(x => x.Id == id) ?? throw new UserError("That ingestion recipe does not exist.");
        var copy = JsonSerializer.Deserialize<IngestRecipe>(JsonSerializer.Serialize(original, Json.Options), Json.Options)!; copy.Folders = copy.Folders.Select(_ => "{Downloads}").ToList(); return copy;
    }
    public IReadOnlyList<LibraryLiteral> Preview(IReadOnlyList<LibrarySelection> items)
    {
        var output = new List<LibraryLiteral>(); foreach (var item in Build(items).Items) Walk(item.Data, "", item, output); return output;
    }
    static void Walk(JsonElement e, string path, LibraryItem item, List<LibraryLiteral> result)
    {
        if (e.ValueKind == JsonValueKind.Object)
        {
            bool predicate = e.TryGetProperty("op", out var op) && op.ValueKind == JsonValueKind.String && e.TryGetProperty("col", out _);
            foreach (var p in e.EnumerateObject()) { string field = path.Length == 0 ? p.Name : path + "." + p.Name; if (predicate && p.Name is "value" or "values" or "a" or "b") { if (p.Value.ValueKind == JsonValueKind.Array) foreach (var v in p.Value.EnumerateArray()) result.Add(new(item.Kind, item.Id, field, v.Clone())); else result.Add(new(item.Kind, item.Id, field, p.Value.Clone())); } else Walk(p.Value, field, item, result); }
        }
        else if (e.ValueKind == JsonValueKind.Array) { int i = 0; foreach (var v in e.EnumerateArray()) Walk(v, path + "[" + i++ + "]", item, result); }
    }
    public void Export(IReadOnlyList<LibrarySelection> selections, string target) => ExternalLibrary.Write(target, JsonSerializer.Serialize(Build(selections), Json.Options));
    public (IReadOnlyList<object> imported, IReadOnlyList<string> renamed) Import(string path)
    {
        var file = JsonSerializer.Deserialize<LibraryFile>(ExternalLibrary.Read(path), Json.Options) ?? throw new UserError("The library file is invalid."); if (file.Format != "exlib1") throw new UserError("The library format is not supported.");
        var imported = new List<object>(); var renamed = new List<string>();
        foreach (var item in file.Items)
        {
            string id = SafeId(item.Id); bool exists(string candidate) => item.Kind switch { "view" => views.List().Any(x => x.Id == candidate), "link" => views.ListLinks().Any(x => x.Id == candidate), "alert" => alerts.List().Any(x => x.Id == candidate), "reportRecipe" => recipes.List().Any(x => x.Id == candidate), "dashboard" => services.Config.Exists("dashboards/" + candidate + ".json"), "ingestRecipe" => services.Ingestion.Recipes.Any(x => x.Id == candidate), _ => throw new UserError("That library item kind is unknown.") };
            string original = id; for (int n = 2; exists(id); n++) id = original + "-" + n; if (id != original) renamed.Add(original + " → " + id);
            string name;
            switch (item.Kind)
            {
                case "view": { var x = item.Data.Deserialize<SavedView>(Json.Options)!; x.Id = id; views.Save(x); name = x.Name; break; }
                case "link": { var x = item.Data.Deserialize<Link>(Json.Options)!; x.Id = id; views.SaveLink(x); name = x.Name; break; }
                case "alert": { var x = item.Data.Deserialize<AlertRule>(Json.Options)!; x.Id = id; alerts.Save(x); name = x.Name; break; }
                case "reportRecipe": { var x = item.Data.Deserialize<ReportRecipe>(Json.Options)!; x.Id = id; recipes.Save(x); name = x.Name; break; }
                case "dashboard": { var x = item.Data.Deserialize<Dashboard>(Json.Options)!; x.Id = id; services.Config.WriteAllText("dashboards/" + id + ".json", JsonSerializer.Serialize(x, Json.Options)); name = x.Name; break; }
                case "ingestRecipe": { var x = item.Data.Deserialize<IngestRecipe>(Json.Options)!; x.Id = id; string downloads = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"); x.Folders = x.Folders.Select(_ => downloads).ToList(); services.Ingestion.SaveRecipe(x); name = x.Name; break; }
                default: throw new UserError("That library item kind is unknown.");
            }
            imported.Add(new { kind = item.Kind, id, name });
        }
        return (imported, renamed);
    }
}
