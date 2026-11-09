using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data.Query;

public sealed class RecipeExport { public string Target { get; set; } = ""; public string? Template { get; set; } public string Sheet { get; set; } = "Sheet1"; public string StartCell { get; set; } = "A1"; }
public sealed class RecipeStep { public string Kind { get; set; } = ""; public QuerySpec? Spec { get; set; } public string? ViewId { get; set; } public RecipeExport? Export { get; set; } public List<PackSheet> Sheets { get; set; } = new(); }
public sealed class PackSheet { public QuerySpec? Spec { get; set; } public string? ViewId { get; set; } public string Sheet { get; set; } = "Sheet1"; public string StartCell { get; set; } = "A1"; }
public sealed class Schedule { public string Kind { get; set; } = ""; public string? Time { get; set; } public int? Day { get; set; } public string? ReportType { get; set; } }
public sealed class ReportRecipe { public string Id { get; set; } = ""; public string Name { get; set; } = ""; public List<RecipeStep> Steps { get; set; } = new(); public Schedule? Schedule { get; set; } public bool Enabled { get; set; } public RecipeRunStatus? LastRun { get; set; } }
public sealed record RecipeRunStatus(DateTime At, bool Ok, string Message);
public sealed record RecipeRun(bool Ok, string Message, string? ResultId = null);
public sealed class ReportRecipes
{
    readonly DataServices services; readonly Views views; readonly object gate = new(); readonly SemaphoreSlim runGate = new(1,1);
    public ReportRecipes(DataServices services, Views views) { this.services = services; this.views = views; }
    static string PathFor(string id) { if (string.IsNullOrWhiteSpace(id) || id.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '-' or '_'))) throw new UserError("That identifier is invalid."); return "reportRecipes/" + id + ".json"; }
    public IReadOnlyList<ReportRecipe> List() => services.Config.List("reportRecipes", "*.json").Select(p => JsonSerializer.Deserialize<ReportRecipe>(services.Config.ReadAllText(p), Json.Options)!).ToArray();
    public ReportRecipe Get(string id) { var p = PathFor(id); return services.Config.Exists(p) ? JsonSerializer.Deserialize<ReportRecipe>(services.Config.ReadAllText(p), Json.Options)! : throw new UserError("That recipe does not exist."); }
    public ReportRecipe Save(ReportRecipe recipe) { lock (gate) { if (services.Config.Exists(PathFor(recipe.Id))) recipe.LastRun = Get(recipe.Id).LastRun; Write(recipe); return recipe; } }
    void Write(ReportRecipe recipe) => services.Config.WriteAllText(PathFor(recipe.Id), JsonSerializer.Serialize(recipe, Json.Options));
    void Record(string id, RecipeRun run)
    {
        lock (gate) { if (!services.Config.Exists(PathFor(id))) return; var recipe = Get(id); recipe.LastRun = new RecipeRunStatus(DateTime.UtcNow, run.Ok, run.Message); Write(recipe); }
    }
    public void Delete(string id) => services.Config.Delete(PathFor(id));
    QuerySpec Spec(QuerySpec? spec, string? viewId) => spec ?? (viewId == null ? throw new UserError("The recipe step needs a query.") : views.Get(viewId).Spec);
    public static string Expand(string target, DateTime date, string period, string name) => target.Replace("{yyyy}", date.ToString("yyyy"), StringComparison.Ordinal).Replace("{MM}", date.ToString("MM"), StringComparison.Ordinal).Replace("{dd}", date.ToString("dd"), StringComparison.Ordinal).Replace("{period}", period, StringComparison.Ordinal).Replace("{name}", name, StringComparison.Ordinal);
    static ExportSheet Sheet(Table table, string sheet, string start)
    {
        var cols = table.Columns.Where(c => !c.Name.StartsWith('_')).ToArray(); var output = new ExportSheet { Sheet = sheet, StartCell = start, Columns = cols.Select(c => c.Name).ToList(), NumberFormats = cols.Select(c => c.Type == ColType.Date ? "yyyy-mm-dd" : "").ToList() };
        for (int r = 0; r < table.Rows; r++) { int row = r; output.Rows.Add(cols.Select(c => c.IsNull(row) ? null : c.Text(row)).ToArray()); }
        return output;
    }
    void Emit(string id, RecipeRun run) { try { services.Emit("data.recipeRun", new { recipeId=id, ok=run.Ok, message=run.Message }); } catch { } }
    public async Task<RecipeRun> Run(string id, string period = "", CancellationToken token = default)
    {
        await runGate.WaitAsync(token).ConfigureAwait(false);
        try
        {
            var recipe = Get(id); Table? last = null; string? resultId = null; var now = DateTime.Now;
            foreach (var step in recipe.Steps)
            {
                token.ThrowIfCancellationRequested();
                if (step.Kind == "query") { var spec = Spec(step.Spec, step.ViewId); last = Engine.Run(spec, services.Tables, token).Table; resultId = services.Results.Add(last, new ResultMeta("query", spec, (spec.GroupBy?.Count ?? 0) > 0 || (spec.Aggregates?.Count ?? 0) > 0, spec.Source)); }
                else if (step.Kind is "export" or "pack")
                {
                    var target = step.Export ?? throw new UserError("The export step needs a target."); var job = new ExportJob { Target = Expand(target.Target, now, period, recipe.Name), Template = target.Template };
                    if (step.Kind == "export") job.Sheets.Add(Sheet(last ?? throw new UserError("Run a query before export."), target.Sheet, target.StartCell));
                    else foreach (var item in step.Sheets) { var table = item.Spec == null && item.ViewId == null ? last ?? throw new UserError("Run a query before packing.") : Engine.Run(Spec(item.Spec, item.ViewId), services.Tables, token).Table; job.Sheets.Add(Sheet(table, item.Sheet, item.StartCell)); }
                    if (job.Sheets.Count == 0) throw new UserError("The pack needs at least one sheet."); await services.Exporter.ExportAsync(job, null, token).ConfigureAwait(false);
                }
                else throw new UserError("The recipe contains an unknown step.");
            }
            var success = new RecipeRun(true, "Recipe completed.", resultId); Record(id, success); Emit(id, success); return success;
        }
        catch (Exception) when (!token.IsCancellationRequested) { var failure = new RecipeRun(false, "The recipe could not be completed."); try { Record(id, failure); } catch (IOException) { } Emit(id, failure); return failure; }
        finally { runGate.Release(); }
    }
}
