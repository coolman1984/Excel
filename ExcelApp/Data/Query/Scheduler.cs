using System.Text.Json;

namespace ExcelApp.Data.Query;

public sealed class Scheduler : IDisposable
{
    readonly DataServices services; readonly ReportRecipes recipes; readonly SemaphoreSlim gate = new(1, 1); readonly System.Collections.Concurrent.ConcurrentDictionary<string, DateTime> last = new(); System.Threading.Timer? timer; bool disposed;
    public Scheduler(DataServices services, ReportRecipes recipes)
    {
        this.services = services; this.recipes = recipes;
        if (services.Config.Exists("scheduler.json")) { var saved = JsonSerializer.Deserialize<Dictionary<string, DateTime>>(services.Config.ReadAllText("scheduler.json"), Json.Options); if (saved != null) foreach (var item in saved) last[item.Key] = item.Value; }
    }
    public void Start() { if (services.Ingestion != null) services.Ingestion.Loaded += OnLoaded; timer = new System.Threading.Timer(_ => _ = Tick(DateTime.Now), null, TimeSpan.Zero, TimeSpan.FromSeconds(30)); }
    void OnLoaded(LoadRecord record) => _ = RunAfterLoad(record);
    async Task RunAfterLoad(LoadRecord record) { await gate.WaitAsync(); try { if (disposed) return; foreach (var r in recipes.List()) if (r.Enabled && r.Schedule?.Kind == "afterLoad" && string.Equals(r.Schedule.ReportType, record.RecipeId, StringComparison.OrdinalIgnoreCase)) { await recipes.Run(r.Id, record.Periods.FirstOrDefault() ?? ""); last[r.Id] = DateTime.Now; Save(); } } finally { gate.Release(); } }
    void Save() => services.Config.WriteAllText("scheduler.json", JsonSerializer.Serialize(last, Json.Options));
    public DateTime? LastRun(string id) => last.TryGetValue(id, out var dt) ? dt : null;
    static DateTime? Due(Schedule schedule, DateTime now)
    {
        if (!TimeSpan.TryParse(schedule.Time, out var time) || time < TimeSpan.Zero || time >= TimeSpan.FromDays(1)) return null;
        var day = now.Date; DateTime due;
        switch (schedule.Kind)
        {
            case "daily": due = day + time; if (due > now) due = due.AddDays(-1); return due;
            case "weekly": int weekday = schedule.Day ?? 1; if (weekday < 1 || weekday > 7) return null; int current = ((int)day.DayOfWeek + 6) % 7 + 1; due = day.AddDays(-((current - weekday + 7) % 7)) + time; if (due > now) due = due.AddDays(-7); return due;
            case "monthly": int monthDay = schedule.Day ?? 1; if (monthDay < 1 || monthDay > 31) return null; due = new DateTime(day.Year, day.Month, Math.Min(monthDay, DateTime.DaysInMonth(day.Year, day.Month))) + time; if (due > now) { var previous = day.AddMonths(-1); due = new DateTime(previous.Year, previous.Month, Math.Min(monthDay, DateTime.DaysInMonth(previous.Year, previous.Month))) + time; } return due;
            default: return null;
        }
    }
    public async Task Tick(DateTime now)
    {
        if (!await gate.WaitAsync(0)) return; try { if (disposed) return; foreach (var r in recipes.List()) { if (!r.Enabled || r.Schedule == null) continue; var due = Due(r.Schedule, now); if (due == null || (last.TryGetValue(r.Id, out var previous) && previous >= due)) continue; await recipes.Run(r.Id); last[r.Id] = now; Save(); } } catch (Exception) { } finally { gate.Release(); }
    }
    public void Dispose() { disposed = true; timer?.Dispose(); if (services.Ingestion != null) services.Ingestion.Loaded -= OnLoaded; }
}
