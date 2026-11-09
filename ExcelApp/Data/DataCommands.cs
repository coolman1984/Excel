using System.Text.Json;
using Microsoft.Win32;
namespace ExcelApp.Data;
public static class DataCommands
{
    static string String(JsonElement e,string name,string fallback="") => e.ValueKind==JsonValueKind.Object && e.TryGetProperty(name,out var v) && v.ValueKind==JsonValueKind.String ? v.GetString() ?? fallback : fallback;
    static bool Bool(JsonElement e,string name,bool fallback=false) => e.ValueKind==JsonValueKind.Object && e.TryGetProperty(name,out var v) && v.ValueKind is JsonValueKind.True or JsonValueKind.False ? v.GetBoolean() : fallback;
    static int Int(JsonElement e,string name,int fallback=0) => e.ValueKind==JsonValueKind.Object && e.TryGetProperty(name,out var v) && v.TryGetInt32(out var n)?n:fallback;
    static T Read<T>(JsonElement e,string name) => e.TryGetProperty(name,out var v)?JsonSerializer.Deserialize<T>(v.GetRawText(),Json.Options) ?? throw new UserError("Missing or invalid data."):throw new UserError("Missing data.");
    public static void Register(Dictionary<string,Func<JsonElement,Task<object?>>> h,DataServices s) {
        var queue=new WorkQueue();var source=new ExcelReader();var catalog=new Catalog(s.Store);var provider=new TableProvider(s.Store,catalog);var ingestion=new Ingestion(s.Store,s.Config,source,provider,catalog,queue,s.Emit);var workspace=new Workspace(s.Store,catalog,source,queue,s.Emit);var history=new History(s.Store,s.Results,id=>ingestion.Recipes.FirstOrDefault(x=>x.Id==id));
        s.Settings=s.Config.Exists("settings.json")?JsonSerializer.Deserialize<Settings>(s.Config.ReadAllText("settings.json"),Json.Options) ?? new Settings():new Settings();s.Settings.AllowPdf=false;s.Settings.MaxParallelExcel=1;
        s.Tables=provider;s.Ingestion=ingestion;s.Exporter=new Exporter(s.Store,queue,()=>s.Settings);s.Owned.Add(queue);s.Owned.Add(ingestion);
        void Sync(string command,Func<JsonElement,object?> action) => h[command]=e=>Task.FromResult(action(e));
        void Async(string command,Func<JsonElement,object?> action) => h[command]=e=>Task.Run(()=>action(e));
        object Info(Table t,string kind) {string id=s.Results.Add(t,new ResultMeta(kind,null,false,null));return new {resultId=id,rows=t.Rows,columns=t.Columns.Where(c=>!c.Name.StartsWith('_') || c.Name=="_status").Select(c=>new ColumnInfo(c.Name,c.Type)).ToArray(),ms=0};}
        object SettingsInfo()=>new {settings=s.Settings,storeFolder=s.Store.Folder,configFolder=s.Config.Folder,pdfApproved=false};
        Sync("data.tables.list",_=>new {tables=provider.List()});
        Async("data.columns",e=>new {columns=provider.Get(Read<TableRef>(e,"source")).Columns.Where(x=>!x.Name.StartsWith('_')).Select(x=>new ColumnInfo(x.Name,x.Type)).ToArray()});
        Sync("data.workspace.projects",_=>new {projects=workspace.Projects()});
        Sync("data.workspace.create",e=>new {id=workspace.Create(String(e,"name"))});
        Sync("data.workspace.rename",e=>{workspace.Rename(String(e,"id"),String(e,"name"));return new {};});
        Sync("data.workspace.delete",e=>{workspace.Delete(String(e,"id"));return new {};});
        Sync("data.workspace.add",e=>{workspace.Add(String(e,"projectId"),String(e,"path"));return new {queued=true};});
        Sync("data.workspace.remove",e=>{workspace.Remove(String(e,"projectId"),String(e,"tableId"));return new {};});
        Sync("data.workspace.refresh",e=>{workspace.Refresh(String(e,"projectId"));return new {};});
        Sync("data.ingest.recipes",_=>new {recipes=ingestion.Recipes});
        Sync("data.ingest.saveRecipe",e=>{var recipe=Read<IngestRecipe>(e,"recipe");ingestion.SaveRecipe(recipe);return new {recipe};});
        Sync("data.ingest.deleteRecipe",e=>{ingestion.DeleteRecipe(String(e,"id"));return new {};});
        h["data.ingest.preview"]=e=>{string path=String(e,"path");var result=new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);queue.Enqueue("preview",token=>{try {s.Emit("data.ingest",new {stage="noticed",file=path,message="Reading the sample through Excel."});var sheets=source.ReadSheets(path,token).Where(sheet=>sheet.Rows>0 && sheet.Cols>0).Select(sheet=>{int header=sheet.FirstRow;var typed=Typing.Type(sheet,header,Array.Empty<TypeFix>());return new {name=sheet.Name,rows=sheet.Rows,cols=sheet.Cols,headerRow=header,header=Enumerable.Range(0,sheet.Cols).Select(c=>sheet.Text(0,c)).ToArray(),sample=Enumerable.Range(1,Math.Min(20,Math.Max(0,sheet.Rows-1))).Select(r=>Enumerable.Range(0,sheet.Cols).Select(c=>sheet.Text(r,c)).ToArray()).ToArray(),guesses=typed.Table.Columns.Where(c=>!c.Name.StartsWith('_')).Select(c=>new {name=c.Name,type=c.Type.ToString().ToLowerInvariant(),fixes=typed.Cleanup.Items.Where(x=>x.Column==c.Name).Select(x=>x.Fix).ToArray()}).ToArray()};}).ToArray();result.TrySetResult(new {sheets});}catch(Exception ex){result.TrySetException(ex);}return Task.CompletedTask;});return result.Task;};
        Sync("data.ingest.status",_=>ingestion.Status());
        Sync("data.ingest.scanNow",_=>{ingestion.ScanNow();return new {};});
        Sync("data.ingest.loadFile",e=>{ingestion.QueueLoad(String(e,"path"),String(e,"recipeId"));return new {queued=true};});
        Async("data.ingest.log",e=>new {records=ingestion.LogRecords().Where(x=>String(e,"recipeId")=="" || x.RecipeId==String(e,"recipeId")).Take(Math.Clamp(Int(e,"limit",100),0,10000)).ToArray()});
        Async("data.ingest.quarantine",_=>new {records=ingestion.LogRecords().Where(x=>x.Status=="quarantined").ToArray()});
        Async("data.ingest.accept",e=>new {record=ingestion.Accept(String(e,"loadId"),Bool(e,"split"),Bool(e,"updateRecipe"))});
        Async("data.ingest.reject",e=>new {record=ingestion.Reject(String(e,"loadId"))});
        Async("data.ingest.rollback",e=>new {record=ingestion.Rollback(String(e,"loadId"))});
        Async("data.ingest.undoCleanup",e=>new {record=ingestion.UndoCleanup(String(e,"loadId"),String(e,"column"),String(e,"fix"))});
        Async("data.history.versions",e=>{var (periods,versions)=history.List(String(e,"reportTypeId"),String(e,"period") is string p && p.Length>0?p:null);return new {periods,versions};});
        Async("data.history.diff",e=>history.Diff(String(e,"reportTypeId"),String(e,"period"),Int(e,"from"),Int(e,"to")));
        Async("data.history.trend",e=>new {points=history.Trend(String(e,"reportTypeId"),String(e,"column"),String(e,"agg"),String(e,"mode"),String(e,"period") is string p && p.Length>0?p:null,Int(e,"last") is int n && n>0?n:null)});
        Async("data.history.lock",e=>{string report=String(e,"reportTypeId");new Versions(s.Store,report).Lock(String(e,"period"),Bool(e,"locked"));provider.Invalidate(report);return new {};});
        Async("data.quality.rows",e=>{string id=String(e,"loadId");var record=ingestion.LogRecords().FirstOrDefault(x=>x.Id==id) ?? throw new UserError("The load was not found.");int index=Int(e,"flag",-1);if(index<0 || record.Quality==null || index>=record.Quality.Flags.Count)throw new UserError("The quality flag was not found.");var table=TableFile.Open(s.Store,$"ingest/quality/{id}.xvt").Load();string indexPath=$"ingest/quality/{id}.{index}.json";var indices=s.Store.Exists(indexPath)?JsonSerializer.Deserialize<int[]>(s.Store.ReadAllText(indexPath),Json.Options) ?? Array.Empty<int>():record.Quality.Flags[index].SampleRows;var rows=indices.Where(i=>i>=0 && i<table.Rows).ToArray();return Info(Ingestion.Slice(table,rows),"quality");});
        Sync("data.settings.get",_=>SettingsInfo());
        Sync("data.settings.set",e=>{var settings=Read<Settings>(e,"settings");if(settings.AllowPdf)throw new UserError("PDF export is not approved.");settings.MaxParallelExcel=1;if(!s.Config.IsTest) {using var key=Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");if(settings.StartWithWindows)key.SetValue("ExcelApp",$"\"{Environment.ProcessPath}\" --tray");else key.DeleteValue("ExcelApp",false);}s.Config.WriteAllText("settings.json",JsonSerializer.Serialize(settings,Json.Options));s.Settings=settings;return SettingsInfo();});
    }
}
