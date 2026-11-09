using System.Text.Json;
using ExcelApp.Data.Query;
using ExcelViewer.Core;

namespace ExcelApp.Data;

public static class QueryCommands
{
    public static void Register(Dictionary<string, Func<JsonElement, Task<object?>>> h, DataServices s)
    {
        var sandboxes = new Sandboxes(s, s.Tables); s.Tables = new SandboxProvider(s.Tables, sandboxes);
        var questions = new Questions(s);
        var views = new Views(s);
        var dashboards = new Dashboards(s); s.Owned.Add(dashboards);
        var alerts = new Alerts(s, views);
        s.Owned.Add(alerts); alerts.Start();
        var recipes = new ReportRecipes(s, views); var library = new Library(s, views, alerts, recipes); var scheduler = new Scheduler(s, recipes); s.Owned.Add(scheduler); scheduler.Start();
        static string Id(JsonElement a) => a.GetProperty("id").GetString()!;
        static T Read<T>(JsonElement a, string name) => JsonSerializer.Deserialize<T>(a.GetProperty(name), Json.Options)!;
        h["data.page"] = a => Task.FromResult<object?>(Format.Page(s.Results, a.GetProperty("resultId").GetString()!, a.GetProperty("from").GetInt32(), a.GetProperty("count").GetInt32()));
        h["data.release"] = a => { s.Results.Release(a.GetProperty("resultId").GetString()!); return Task.FromResult<object?>(new { }); };
        h["data.run"] = a => Task.Run<object?>(() => Drill.Run(Read<QuerySpec>(a, "spec"), s));
        h["data.drill"] = a => Task.Run<object?>(() => Drill.Open(s, a.GetProperty("resultId").GetString()!, a.GetProperty("row").GetInt32()));
        h["data.lineage"] = a => Task.FromResult<object?>(Drill.Lineage(s, a.GetProperty("resultId").GetString()!, a.GetProperty("row").GetInt32(), a.GetProperty("col").GetString()!));
        h["data.export"] = async a =>
        {
            string resultId = a.GetProperty("resultId").GetString()!; if (!s.Results.TryGet(resultId, out var table, out _)) throw new UserError("That result is no longer available.");
            var cols = table.Columns.Where(c => !c.Name.StartsWith('_')).ToArray();
            var sheet = new ExportSheet { Sheet = a.TryGetProperty("sheet", out var sheetName) ? sheetName.GetString() ?? "Sheet1" : "Sheet1", StartCell = a.TryGetProperty("startCell", out var cell) ? cell.GetString() ?? "A1" : "A1", Columns = cols.Select(c => c.Name).ToList(), NumberFormats = cols.Select(c => c.Type == ColType.Date ? "yyyy-mm-dd" : "").ToList() };
            for (int r = 0; r < table.Rows; r++) { int row = r; sheet.Rows.Add(cols.Select(c => c.IsNull(row) ? null : c.Text(row)).ToArray()); }
            var job = new ExportJob { Target = a.GetProperty("target").GetString()!, Template = a.TryGetProperty("template", out var template) ? template.GetString() : null, Pdf = a.TryGetProperty("pdf", out var pdf) && pdf.GetBoolean(), Sheets = new() { sheet } };
            if (job.Pdf && !s.Settings.AllowPdf) throw new UserError("PDF export is not enabled.");
            await s.Exporter.ExportAsync(job, null, CancellationToken.None); return new { ok = true };
        };
        h["data.views.list"] = _ => Task.FromResult<object?>(new { views = views.Status().Select(v => new { v.View.Id, v.View.Name, v.View.Spec, v.View.Exception, v.View.Folder, v.Count, v.RefreshedAt }).ToArray() });
        h["data.views.save"] = a => Task.Run<object?>(() => { var view = views.Save(Read<SavedView>(a, "view")); views.Refresh(); return new { view }; });
        h["data.views.delete"] = a => { views.Delete(Id(a)); return Task.FromResult<object?>(new { }); };
        h["data.views.run"] = a => Task.Run<object?>(() => Drill.Run(views.Get(Id(a)).Spec, s));
        h["data.links.list"] = _ => Task.FromResult<object?>(new { links = views.ListLinks() });
        h["data.links.save"] = a => Task.FromResult<object?>(new { link = views.SaveLink(Read<Link>(a, "link")) });
        h["data.links.delete"] = a => { views.DeleteLink(Id(a)); return Task.FromResult<object?>(new { }); };
        h["data.links.run"] = a => Task.Run<object?>(() => Drill.Run(views.LinkSpec(views.GetLink(Id(a))), s));
        h["data.alerts.list"] = _ => Task.FromResult<object?>(new { alerts = alerts.List() });
        h["data.alerts.save"] = a => Task.FromResult<object?>(new { alert = alerts.Save(Read<AlertRule>(a, "alert")) });
        h["data.alerts.delete"] = a => { alerts.Delete(Id(a)); return Task.FromResult<object?>(new { }); };
        h["data.alerts.fired"] = a => Task.FromResult<object?>(new { events = alerts.Fired(a.TryGetProperty("limit", out var limit) ? limit.GetInt32() : 100) });
        h["data.recipes.list"] = _ => Task.FromResult<object?>(new { recipes = recipes.List().Select(r => new { r.Id, r.Name, r.Steps, r.Schedule, r.Enabled, lastRun = r.LastRun?.At, lastStatus = r.LastRun == null ? null : new { r.LastRun.Ok, r.LastRun.Message } }).ToArray() });
        h["data.recipes.save"] = a => Task.FromResult<object?>(new { recipe = recipes.Save(Read<ReportRecipe>(a, "recipe")) });
        h["data.recipes.delete"] = a => { recipes.Delete(Id(a)); return Task.FromResult<object?>(new { }); };
        h["data.recipes.run"] = async a => await recipes.Run(Id(a));
        h["data.sandbox.create"] = a => Task.Run<object?>(() => { var x = sandboxes.Create(Read<TableRef>(a, "base"), a.GetProperty("name").GetString()!); return new { x.Id, x.Name, x.Base, x.Edits, x.CreatedAt }; });
        h["data.sandbox.list"] = _ => Task.FromResult<object?>(new { sandboxes = sandboxes.List().Select(x => new { x.Id, x.Name, x.Base, x.Edits, x.CreatedAt }).ToArray() });
        h["data.sandbox.discard"] = a => { sandboxes.Discard(Id(a)); return Task.FromResult<object?>(new { }); };
        h["data.sandbox.edit"] = a => Task.Run<object?>(() => new { edits = sandboxes.Edit(Id(a), Read<List<SandboxEdit>>(a, "edits")) });
        h["data.sandbox.compare"] = a => Task.Run<object?>(() => sandboxes.Compare(Id(a)));
        h["data.sandbox.keep"] = a => Task.Run<object?>(() => new { version = sandboxes.Keep(Id(a), a.GetProperty("reportTypeId").GetString()!, a.GetProperty("period").GetString()!, a.GetProperty("note").GetString()!) });
        h["data.library.preview"] = a => Task.Run<object?>(() => new { literals = library.Preview(Read<List<LibrarySelection>>(a, "items")) });
        h["data.library.export"] = a => Task.Run<object?>(() => { library.Export(Read<List<LibrarySelection>>(a, "items"), a.GetProperty("target").GetString()!); return new { ok = true }; });
        h["data.library.import"] = a => Task.Run<object?>(() => { var result = library.Import(a.GetProperty("path").GetString()!); return new { result.imported, result.renamed }; });
        h["data.dashboard.list"] = _ => Task.FromResult<object?>(new { dashboards = dashboards.List() });
        h["data.dashboard.save"] = a => Task.FromResult<object?>(new { dashboard = dashboards.Save(Read<Dashboard>(a, "dashboard")) });
        h["data.dashboard.delete"] = a => { dashboards.Delete(Id(a)); return Task.FromResult<object?>(new { }); };
        h["data.dashboard.values"] = a => Task.Run<object?>(() => new { values = dashboards.Values(Id(a)) });
        h["data.ask"] = a => Task.Run<object?>(() => questions.Ask(a.GetProperty("text").GetString()!));
        if (s.Ingestion != null) s.Ingestion.Loaded += record => { _ = Task.Run(() => { try { views.Refresh(record.RecipeId); } catch (Exception) { } }); };
        _ = Task.Run(() => { try { views.Refresh(); } catch (Exception) { } });
    }
}
