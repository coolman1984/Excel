using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data.Query;

public sealed class SavedView
{
    public string Id { get; set; } = ""; public string Name { get; set; } = ""; public QuerySpec Spec { get; set; } = new(); public bool Exception { get; set; } public string Folder { get; set; } = "";
}
public sealed class Link
{
    public string Id { get; set; } = ""; public string Name { get; set; } = ""; public TableRef Left { get; set; } = new("table", ""); public TableRef Right { get; set; } = new("table", ""); public List<JoinKey> On { get; set; } = new(); public List<LinkCompare> Compare { get; set; } = new(); public string Fuzzy { get; set; } = "none";
}
public sealed record LinkCompare(string Left, string Right, string As);
public sealed record ViewStatus(SavedView View, int Count, DateTime? RefreshedAt);

public sealed class Views
{
    readonly DataServices services; readonly object gate = new(); readonly Dictionary<string, (int count, DateTime at)> counts = new();
    public Views(DataServices services) => this.services = services;
    static string PathFor(string folder, string id)
    {
        if (string.IsNullOrWhiteSpace(id) || id.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '-' or '_'))) throw new UserError("That identifier is invalid.");
        return folder + "/" + id + ".json";
    }
    public IReadOnlyList<SavedView> List() => services.Config.List("views", "*.json").Select(p => JsonSerializer.Deserialize<SavedView>(services.Config.ReadAllText(p), Json.Options)!).ToArray();
    public SavedView Get(string id) { var path = PathFor("views", id); return services.Config.Exists(path) ? JsonSerializer.Deserialize<SavedView>(services.Config.ReadAllText(path), Json.Options)! : throw new UserError("That view does not exist."); }
    public SavedView Save(SavedView view) { services.Config.WriteAllText(PathFor("views", view.Id), JsonSerializer.Serialize(view, Json.Options)); lock (gate) counts.Remove(view.Id); return view; }
    public void Delete(string id) { services.Config.Delete(PathFor("views", id)); lock (gate) counts.Remove(id); }
    public IReadOnlyList<ViewStatus> Status() => List().Select(v => { lock (gate) return counts.TryGetValue(v.Id, out var c) ? new ViewStatus(v, c.count, c.at) : new ViewStatus(v, 0, null); }).ToArray();
    public void Refresh(string? reportType = null)
    {
        foreach (var view in List()) if (view.Exception && (reportType == null || Depends(view.Spec, reportType)))
        {
            var result = Engine.Run(view.Spec, services.Tables, CancellationToken.None).Table; lock (gate) counts[view.Id] = (result.Rows, DateTime.UtcNow);
        }
    }
    public static bool Depends(QuerySpec spec, string reportType) => (spec.Source.Kind == "report" && spec.Source.Id == reportType) || (spec.Joins?.Any(j => j.Right.Kind == "report" && j.Right.Id == reportType) ?? false);
    public IReadOnlyList<Link> ListLinks() => services.Config.List("links", "*.json").Select(p => JsonSerializer.Deserialize<Link>(services.Config.ReadAllText(p), Json.Options)!).ToArray();
    public Link GetLink(string id) { string p = PathFor("links", id); return services.Config.Exists(p) ? JsonSerializer.Deserialize<Link>(services.Config.ReadAllText(p), Json.Options)! : throw new UserError("That link does not exist."); }
    public Link SaveLink(Link link) { services.Config.WriteAllText(PathFor("links", link.Id), JsonSerializer.Serialize(link, Json.Options)); return link; }
    public void DeleteLink(string id) => services.Config.Delete(PathFor("links", id));
    public QuerySpec LinkSpec(Link link)
    {
        var spec = new QuerySpec { Source = link.Left, Joins = new() { new JoinSpec { Right = link.Right, On = link.On, Kind = "left", Fuzzy = link.Fuzzy, Prefix = "right." } }, Computed = new() };
        foreach (var c in link.Compare) { string l = "[" + c.Left + "]", r = "[right." + c.Right + "]"; spec.Computed.Add(new Computed(c.As, l + " - " + r)); spec.Computed.Add(new Computed(c.As + " %", "pct(" + l + "," + r + ")")); }
        return spec;
    }
}
