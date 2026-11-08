using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data.Query;

public sealed class AlertRule
{
    public string Id { get; set; } = ""; public string Name { get; set; } = ""; public string? ViewId { get; set; } public QuerySpec? Spec { get; set; } public string Measure { get; set; } = "rows"; public string Op { get; set; } = "gt"; public double Value { get; set; } public List<string> ReportTypes { get; set; } = new(); public bool Enabled { get; set; }
}
public sealed record AlertEvent(string Id, string RuleId, string Name, DateTime At, double Value, string Message);
public sealed class Alerts : IDisposable
{
    readonly DataServices services; readonly Views views; readonly object gate = new(); readonly Dictionary<string, double> last = new(); readonly Dictionary<string, double> fired = new(); System.Threading.Timer? timer;
    sealed class State { public Dictionary<string, double> Last { get; set; } = new(); public Dictionary<string, double> Fired { get; set; } = new(); }
    public Alerts(DataServices services, Views views)
    {
        this.services = services; this.views = views;
        if (services.Store.Exists("alerts/state.json"))
        {
            var state = JsonSerializer.Deserialize<State>(services.Store.ReadAllText("alerts/state.json"), Json.Options);
            if (state != null) { foreach (var pair in state.Last) last[pair.Key] = pair.Value; foreach (var pair in state.Fired) fired[pair.Key] = pair.Value; }
        }
    }
    void SaveState() => services.Store.WriteAllText("alerts/state.json", JsonSerializer.Serialize(new State { Last = last, Fired = fired }, Json.Options));
    static string PathFor(string id) { if (string.IsNullOrWhiteSpace(id) || id.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '-' or '_'))) throw new UserError("That identifier is invalid."); return "alerts/" + id + ".json"; }
    public IReadOnlyList<AlertRule> List() => services.Config.List("alerts", "*.json").Select(p => JsonSerializer.Deserialize<AlertRule>(services.Config.ReadAllText(p), Json.Options)!).ToArray();
    public AlertRule Save(AlertRule rule) { services.Config.WriteAllText(PathFor(rule.Id), JsonSerializer.Serialize(rule, Json.Options)); lock (gate) { last.Remove(rule.Id); fired.Remove(rule.Id); SaveState(); } return rule; }
    public void Delete(string id) { services.Config.Delete(PathFor(id)); lock (gate) { last.Remove(id); fired.Remove(id); SaveState(); } }
    public IReadOnlyList<AlertEvent> Fired(int limit = 100)
    {
        if (limit < 0) throw new UserError("The alert limit is invalid.");
        if (!services.Store.Exists("alerts/log.jsonl")) return Array.Empty<AlertEvent>();
        using var reader = new StreamReader(services.Store.OpenRead("alerts/log.jsonl")); var events = new Queue<AlertEvent>(); string? line;
        while ((line = reader.ReadLine()) != null) { var item = JsonSerializer.Deserialize<AlertEvent>(line, Json.Options); if (item == null || limit == 0) continue; events.Enqueue(item); if (events.Count > limit) events.Dequeue(); }
        return events.Reverse().ToArray();
    }
    public void Start()
    {
        if (services.Ingestion != null) services.Ingestion.Loaded += OnLoaded;
        timer = new System.Threading.Timer(_ => { try { Evaluate(null); } catch (Exception) { } }, null, TimeSpan.FromHours(1), TimeSpan.FromHours(1));
    }
    void OnLoaded(LoadRecord record) => _ = Task.Run(() => Evaluate(record.RecipeId));
    public void Evaluate(string? reportType)
    {
        lock (gate)
        {
            foreach (var rule in List())
            {
                if (!rule.Enabled || (reportType != null && rule.ReportTypes.Count > 0 && !rule.ReportTypes.Contains(reportType, StringComparer.OrdinalIgnoreCase))) continue;
                try { EvaluateRule(rule); SaveState(); } catch (Exception) { /* A broken rule must not stop other rules. */ }
            }
        }
    }
    void EvaluateRule(AlertRule rule)
    {
        var spec = rule.Spec ?? (rule.ViewId != null ? views.Get(rule.ViewId).Spec : throw new UserError("The alert has no query."));
        var result = Engine.Run(spec, services.Tables, CancellationToken.None).Table;
        double value = rule.Measure == "rows" ? result.Rows : result.Rows > 0 && result[rule.Measure] is Column c && !c.IsNull(0) && c.Type != ColType.Text ? c.Num(0) : throw new UserError("The alert measure is unavailable.");
        bool hadLast = last.TryGetValue(rule.Id, out double previous); last[rule.Id] = value;
        bool condition = rule.Op switch { "gt" => value > rule.Value, "ge" => value >= rule.Value, "lt" => value < rule.Value, "le" => value <= rule.Value, "eq" => value == rule.Value, "ne" => value != rule.Value, "changePct" => hadLast && previous != 0 && Math.Abs((value - previous) / previous * 100) >= rule.Value, _ => false };
        if (!condition || (fired.TryGetValue(rule.Id, out double old) && old == value)) return;
        var alert = new AlertEvent(Guid.NewGuid().ToString("N"), rule.Id, rule.Name, DateTime.UtcNow, value, "Alert condition met.");
        services.Store.AppendLine("alerts/log.jsonl", JsonSerializer.Serialize(alert, new JsonSerializerOptions(Json.Options) { WriteIndented = false })); fired[rule.Id] = value; services.Emit("data.alert", new { alert });
    }
    public void Dispose() { timer?.Dispose(); if (services.Ingestion != null) services.Ingestion.Loaded -= OnLoaded; }
}

