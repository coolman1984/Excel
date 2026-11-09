using ExcelApp.Data;
using ExcelApp.Data.Query;
using System.Text.Json;

public static class AutomationTests
{
    public static void Run(Action<string, bool, string> check, bool offline)
    {
        string suffix = Guid.NewGuid().ToString("N");
        var table = new Table("synthetic", new Column[] { TextColumn.From("Category", new string?[] { "A", "A", "B" }), new NumberColumn("Amount", new[] { 10d, 20d, 5d }, null), new NumberColumn("_v", new[] { 7d, 7d, 7d }, null), new NumberColumn("_r", new[] { 2d, 3d, 4d }, null) });
        var provider = new AutomationProvider(table);
        var ingestion = new FakeIngestion(); var exporter = new FakeExporter(); using var firedSignal = new ManualResetEventSlim();
        var services = new DataServices { Ingestion = ingestion, Exporter = exporter, Tables = provider, Store = StoreRoot.ForTest(@"D:\ExcelAppTest\syn-b5-store-" + suffix), Config = ConfigRoot.ForTest(@"D:\ExcelAppTest\syn-b5-config-" + suffix), Emit = (name, _) => { if (name == "data.alert") firedSignal.Set(); } };
        var views = new Views(services);
        using var filter = JsonDocument.Parse("{\"op\":\"ge\",\"col\":\"Amount\",\"value\":10}");
        var grouped = new QuerySpec { Source = new("report", "r"), Where = filter.RootElement.Clone(), GroupBy = new() { "Category" }, Aggregates = new() { new("sum", "Amount", "Sum") } };
        views.Save(new SavedView { Id = "exceptions", Name = "Exceptions", Spec = grouped, Exception = true }); views.Refresh("r");
        check("exception view refresh", views.Status().Single().Count == 1 && views.Status().Single().RefreshedAt != null, "");
        var result = Drill.Run(grouped, services); var detail = Drill.Open(services, result.ResultId, 0);
        check("drill retains filter", detail.Rows == 2, "");
        var lineage = Drill.Lineage(services, detail.ResultId, 0, "Amount");
        check("lineage source address", lineage.Cell == "Synthetic!AB2" && lineage.Version == 3 && lineage.Value == "10", "");
        var link = new Link { Id = "link", Left = new("report", "r"), Right = new("report", "r"), On = new() { new("Category", "Category") }, Compare = new() { new("Amount", "Amount", "Difference") } };
        views.SaveLink(link); var linked = Engine.Run(views.LinkSpec(views.GetLink("link")), provider, CancellationToken.None).Table;
        check("link difference and percent", linked.Rows == 5 && linked["Difference"] != null && linked["Difference %"] != null, "");
        using var alerts = new Alerts(services, views); alerts.Start();
        alerts.Save(new AlertRule { Id = "rows", Name = "Rows", Spec = new QuerySpec { Source = new("report", "r") }, Measure = "rows", Op = "gt", Value = 2, Enabled = true, ReportTypes = new() { "r" } });
        ingestion.Raise(new LoadRecord { RecipeId = "r" });
        check("alert fires within one load evaluation", firedSignal.Wait(TimeSpan.FromSeconds(5)) && alerts.Fired().Count == 1, "");
        alerts.Evaluate("r"); check("alert suppresses repeated value", alerts.Fired().Count == 1, "");
        alerts.Delete("rows");
        var recipes = new ReportRecipes(services, views);
        recipes.Save(new ReportRecipe { Id = "monthly", Name = "Monthly", Steps = new() { new RecipeStep { Kind = "query", Spec = new QuerySpec { Source = new("report", "r") } }, new RecipeStep { Kind = "export", Export = new RecipeExport { Target = "{period}-{name}.xlsx", Sheet = "Data" } } } });
        var october = recipes.Run("monthly", "2026-10").GetAwaiter().GetResult();
        provider.Set(new Table("synthetic", new Column[] { TextColumn.From("Category", new string?[] { "N" }), new NumberColumn("Amount", new[] { 42d }, null) }));
        var november = recipes.Run("monthly", "2026-11").GetAwaiter().GetResult();
        var persistedRecipe = JsonSerializer.SerializeToElement(new ReportRecipes(services, views).Get("monthly"), Json.Options);
        check("review recipe persists last run status", persistedRecipe.TryGetProperty("lastRun", out var lastRun) && lastRun.ValueKind == JsonValueKind.Object && lastRun.GetProperty("ok").GetBoolean() && lastRun.TryGetProperty("at", out _), "");
        check("recipe replay october november", october.Ok && november.Ok && exporter.Jobs.Count == 2 && exporter.Jobs[0].Sheets[0].Rows.Count == 3 && exporter.Jobs[1].Sheets[0].Rows.Count == 1 && exporter.Jobs[1].Sheets[0].Rows[0][1] == "42" && exporter.Jobs[1].Target == "2026-11-Monthly.xlsx", "");
        recipes.Save(new ReportRecipe { Id = "scheduled", Name = "Scheduled", Enabled = true, Schedule = new Schedule { Kind = "daily", Time = "09:00" }, Steps = new() { new RecipeStep { Kind = "query", Spec = new QuerySpec { Source = new("report", "r") } } } });
        using (var scheduler = new Scheduler(services, recipes)) { var clock = new DateTime(2026, 11, 5, 10, 0, 0); scheduler.Tick(clock).GetAwaiter().GetResult(); var first = scheduler.LastRun("scheduled"); scheduler.Tick(clock.AddMinutes(1)).GetAwaiter().GetResult(); check("scheduler missed run once", first == clock && scheduler.LastRun("scheduled") == clock, ""); }
        using (var scheduler = new Scheduler(services, recipes)) check("scheduler persists last run", scheduler.LastRun("scheduled") == new DateTime(2026, 11, 5, 10, 0, 0), "");
        recipes.Delete("scheduled"); recipes.Delete("monthly"); provider.Set(table);
        var sandboxes = new Sandboxes(services, provider); var sandbox = sandboxes.Create(new("report", "r"), "Synthetic");
        using (var edit = JsonDocument.Parse("99")) sandboxes.Edit(sandbox.Id, new[] { new SandboxEdit { Row = 0, Col = "Amount", Value = edit.RootElement } });
        using (var savedSandbox = JsonDocument.Parse(services.Store.ReadAllText("sandboxes/" + sandbox.Id + ".json")))
            check("review sandbox JSON has no cell values", !savedSandbox.RootElement.TryGetProperty("changes", out var changes) || changes.EnumerateArray().All(e => !e.TryGetProperty("value", out _)), "");
        check("review sandbox edited columns use xvt", services.Store.Exists("sandboxes/" + sandbox.Id + ".xvt"), "");
        check("sandbox report base pinned", sandbox.Base.AsOf != null, "");
        var changed = sandboxes.Get(sandbox.Id); var difference = sandboxes.Compare(sandbox.Id);
        check("sandbox copy on write and compare", changed["Amount"]!.Num(0) == 99 && table["Amount"]!.Num(0) == 10 && ReferenceEquals(changed["Category"], table["Category"]) && difference.Changed == 1 && difference.Columns.Any(c => c.Name == "_status") && Format.Page(services.Results, difference.ResultId, 0, 10).Rows[0].Contains("changed"), "");
        check("sandbox keep", sandboxes.Keep(sandbox.Id, "r", "2026-11", "test") == 4 && provider.Saved!["Amount"]!.Num(0) == 99, ""); sandboxes.Discard(sandbox.Id);
        var millionTable = new Table("synthetic", new Column[] { new NumberColumn("N", new double[1_000_000], null) }); provider.Set(millionTable);
        var sw = System.Diagnostics.Stopwatch.StartNew(); var bigSandbox = sandboxes.Create(new("report", "r"), "Synthetic"); sandboxes.Discard(bigSandbox.Id); sw.Stop();
        Console.WriteLine($"SPEED sandbox create/discard 1M: {sw.Elapsed.TotalSeconds:F3}s (target <5s)"); check("speed sandbox 1M", sw.Elapsed.TotalSeconds < 5, ""); provider.Set(table);
        var library = new Library(services, views, alerts, recipes);
        using var literal = JsonDocument.Parse("{\"op\":\"eq\",\"col\":\"Category\",\"value\":\"filter-literal\"}");
        views.Save(new SavedView { Id = "shared", Name = "Shared", Spec = new QuerySpec { Source = new("report", "r"), Where = literal.RootElement.Clone() } });
        provider.Set(new Table("synthetic", new Column[] { TextColumn.From("Category", new string?[] { "cell-secret-marker" }) }));
        var selection = new[] { new LibrarySelection("view", "shared") }; var preview = library.Preview(selection);
        string libraryPath = services.Config.PathOf("synthetic.exlib"); library.Export(selection, libraryPath); string exported = ExternalLibrary.Read(libraryPath);
        check("library previews literals and excludes table data", preview.Count == 1 && preview[0].Value!.ToString() == "filter-literal" && exported.Contains("filter-literal") && !exported.Contains("cell-secret-marker"), "");
        var imported = library.Import(libraryPath); check("library conflict renamed", imported.imported.Count == 1 && imported.renamed.Count == 1 && views.Get("shared-2").Name == "Shared", "");
        views.Delete("shared"); views.Delete("shared-2"); services.Config.Delete("synthetic.exlib"); provider.Set(table);
        using (var dashboards = new Dashboards(services))
        {
            dashboards.Save(new Dashboard { Id = "home", Name = "Home", Tiles = new() { new Tile { Id = "count", Title = "Count", Spec = new QuerySpec { Source = new("report", "r") }, Measure = "rows" } } });
            check("dashboard scalar and cache", dashboards.Values("home").Single().Value == 3 && ReferenceEquals(dashboards.Values("home"), dashboards.Values("home")), "");
            dashboards.Delete("home");
        }
        var questionTable = new Table("synthetic", new Column[] { TextColumn.From("Category", new string?[] { "A" }), new NumberColumn("Amount", new[] { 10d }, null), new NumberColumn("Date", new[] { DateTime.Today.ToOADate() }, null, ColType.Date) }); provider.Set(questionTable);
        var questions = new Questions(services); string[] examples = { "top 3 Category by Amount", "bottom 3 Category by Amount", "total Amount by Category", "count by Category", "average Amount by Category", "max Amount", "min Amount", "rows where Amount > 2", "Category is A", "show Sales", "duplicates in Category", "blank Category", "sort by Amount desc", "Amount per Category in last month", "trend of Amount by month", "compare Sales and Budget by Category", "how many rows in Sales", "last 7 days of Sales", "Amount this month vs last month" };
        foreach (string example in examples) check("question " + example, questions.Ask(example).Ok, "");
        check("review QuerySpec matches declined history contract", typeof(QuerySpec).GetProperty("History") == null, "");
        var historyQuestion = questions.Ask("what changed in October 2026"); check("question history diff adapter", historyQuestion.Ok && historyQuestion.Spec == null && historyQuestion.Pattern == "history diff|r|2026-10", "");
        check("question unknown", !questions.Ask("unrelated nonsense").Ok, ""); provider.Set(table);
        var handlers = new Dictionary<string, Func<JsonElement, Task<object?>>>(); QueryCommands.Register(handlers, services);
        string[] commands = { "data.run", "data.page", "data.release", "data.drill", "data.lineage", "data.export", "data.views.list", "data.views.save", "data.views.delete", "data.views.run", "data.links.list", "data.links.save", "data.links.delete", "data.links.run", "data.alerts.list", "data.alerts.save", "data.alerts.delete", "data.alerts.fired", "data.recipes.list", "data.recipes.save", "data.recipes.delete", "data.recipes.run", "data.sandbox.create", "data.sandbox.list", "data.sandbox.discard", "data.sandbox.edit", "data.sandbox.compare", "data.sandbox.keep", "data.library.preview", "data.library.export", "data.library.import", "data.dashboard.list", "data.dashboard.save", "data.dashboard.delete", "data.dashboard.values", "data.ask" };
        check("all lane B commands registered", commands.All(handlers.ContainsKey), ""); foreach (var owned in services.Owned) owned.Dispose(); services.Owned.Clear();
        views.Delete("exceptions"); views.DeleteLink("link"); check("delete view and link", views.List().Count == 0 && views.ListLinks().Count == 0, "");
    }
    sealed class FakeExporter : IExporter
    {
        public List<ExportJob> Jobs { get; } = new();
        public Task ExportAsync(ExportJob job, IProgress<string>? progress, CancellationToken token) { Jobs.Add(job); return Task.CompletedTask; }
    }
    sealed class FakeIngestion : IIngestion
    {
        public event Action<LoadRecord>? Loaded;
        public IReadOnlyList<IngestRecipe> Recipes => Array.Empty<IngestRecipe>();
        public void Raise(LoadRecord record) => Loaded?.Invoke(record);
        public void SaveRecipe(IngestRecipe r) { }
        public void DeleteRecipe(string id) { }
        public void Start() { }
        public void Stop() { }
    }
    sealed class AutomationProvider : ITableProvider
    {
        Table table; public AutomationProvider(Table table) => this.table = table;
        public void Set(Table value) => table = value;
        public Table? Saved { get; private set; }
        public Table Get(TableRef r) => table;
        public IReadOnlyList<TableInfo> List() => new[] { new TableInfo("report", "r", "Sales", "test", table.Rows, table.Columns.Select(c => new ColumnInfo(c.Name, c.Type)).ToArray(), DateTime.Now), new TableInfo("report", "b", "Budget", "test", table.Rows, table.Columns.Select(c => new ColumnInfo(c.Name, c.Type)).ToArray(), DateTime.Now) };
        public SourceInfo? SourceOf(int versionId) => versionId == 7 ? new SourceInfo(7, "report", "r", "2026-01", 3, "synthetic.xlsx", "Synthetic", 1, new DateTime(2026, 1, 1), "l", new Dictionary<string, int> { ["Amount"] = 28 }) : null;
        public int SaveAsVersion(string reportTypeId, string period, Table data, string note) { Saved = data; return 4; }
    }
}
