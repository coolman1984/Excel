using ExcelApp.Data;
using ExcelViewer.Core;
public static class IngestTests
{
    static RawSheet Sheet(params string[][] rows) {
        int n=rows.Length, m=rows[0].Length; var k=new byte[n,m]; var d=new double[n,m]; var t=new string[n,m]; var f=new bool[n,m];
        for(int r=0;r<n;r++) for(int c=0;c<m;c++) {t[r,c]=rows[r][c]; k[r,c]=t[r,c].Length==0?CellKind.Empty:CellKind.Text;}
        return new RawSheet(k,d,t,f) { Name="Synthetic" };
    }
    sealed class FakeSource : IRawSource { public Dictionary<string,RawSheet> Sheets=new(); public RawSheet ReadSheet(string file,string? sheet,CancellationToken token)=>Sheets[file]; }
    static void Release(object? value) {if(value!=null && System.Runtime.InteropServices.Marshal.IsComObject(value)) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);}
    static void MakeWorkbook(string path,double amount) {
        dynamic? xl=null,book=null,sheet=null,range=null,dateRange=null;
        try {var type=Type.GetTypeFromProgID("Excel.Application") ?? throw new Exception("Excel unavailable");xl=Activator.CreateInstance(type);xl.Visible=false;xl.DisplayAlerts=false;xl.EnableEvents=false;xl.AutomationSecurity=3;book=xl.Workbooks.Add();sheet=book.Worksheets[1];sheet.Name="Synthetic";
            object[,] values={{"id","date","value"},{"a",new DateTime(2026,10,1).ToOADate(),amount},{"b",new DateTime(2026,10,2).ToOADate(),20.0}};range=sheet.Range["A1:C3"];range.Value2=values;dateRange=sheet.Range["B2:B3"];dateRange.NumberFormat="yyyy-mm-dd";book.SaveAs(path,FileFormat:51);
        } finally {try {if(book!=null)book.Close(false);}catch(Exception){}try {if(xl!=null)xl.Quit();}catch(Exception){}Release(dateRange);Release(range);Release(sheet);Release(book);Release(xl);}
    }
    static void Live(Action<string,bool,string> check) {
        string baseDir=@"D:\ExcelAppTest\syn-ingest";string inDir=Path.Combine(baseDir,"in"),outDir=Path.Combine(baseDir,"out");Directory.CreateDirectory(inDir);Directory.CreateDirectory(outDir);
        string first=Path.Combine(inDir,"syn-out-2026-10.xlsx"),corrected=Path.Combine(inDir,"syn-out-2026-10-corrected.xlsx"),target=Path.Combine(outDir,"syn-pack.xlsx");MakeWorkbook(first,10);MakeWorkbook(corrected,11);
        var store=StoreRoot.ForTest(baseDir);var config=ConfigRoot.ForTest(@"D:\ExcelAppTest\syn-ingest-config");var catalog=new Catalog(store);var provider=new TableProvider(store,catalog);using var queue=new WorkQueue();using var engine=new Ingestion(store,config,new ExcelReader(),provider,catalog,queue,(name,payload)=>{});
        string recipeId="live-"+Guid.NewGuid().ToString("N");engine.SaveRecipe(new IngestRecipe {Id=recipeId,Name="Synthetic",Folders=new(){inDir},Patterns=new(){"*.xlsx"},Sheet="Synthetic",HeaderRow=1,HeaderSignature=new(){"id","date","value"},DateColumn="date",KeyColumns=new(){"id"}});
        var one=engine.Load(first,recipeId);var two=engine.Load(corrected,recipeId);check("live ingestion versions",one.Status=="loaded" && two.Status=="loaded" && two.Versions.GetValueOrDefault("2026-10")==2 && two.Changed==1 && new Versions(store,recipeId).Snapshot().Periods["2026-10"].Versions.First().Status=="archived","");
        var exporter=new Exporter(store,queue,()=>new Settings());var job=new ExportJob {Target=target,Sheets=new(){new ExportSheet {Sheet="Summary",StartCell="A1",Headers=true,Columns=new(){"date","amount"},NumberFormats=new(){"yyyy-mm-dd","0.00"},Rows=new(){new string?[]{"2026-10-01","42.5"}}}}};exporter.ExportAsync(job,null,CancellationToken.None).GetAwaiter().GetResult();
        dynamic? xl=null,book=null,sheet=null,date=null,amount=null;try {var type=Type.GetTypeFromProgID("Excel.Application") ?? throw new Exception("Excel unavailable");xl=Activator.CreateInstance(type);xl.Visible=false;xl.DisplayAlerts=false;xl.EnableEvents=false;xl.AutomationSecurity=3;book=xl.Workbooks.Open(target,UpdateLinks:0,ReadOnly:true,Notify:false);sheet=book.Worksheets["Summary"];date=sheet.Range["A2"];amount=sheet.Range["B2"];check("live export values and formats",Math.Abs(Convert.ToDouble(date.Value2)-new DateTime(2026,10,1).ToOADate())<1e-9 && Math.Abs(Convert.ToDouble(amount.Value2)-42.5)<1e-9 && ((string)date.NumberFormat).Contains("yyyy",StringComparison.OrdinalIgnoreCase) && ((string)amount.NumberFormat).Contains("0.00",StringComparison.Ordinal),"");}
        finally {try {if(book!=null)book.Close(false);}catch(Exception){}try {if(xl!=null)xl.Quit();}catch(Exception){}Release(amount);Release(date);Release(sheet);Release(book);Release(xl);}
    }
    public static void Run(Action<string,bool,string> check, bool offline) {
        try {
        var s=Sheet(new[]{"code","amount","word"},new[]{"001","1,250"," Abc "},new[]{"002","25%","DeF"});
        var r=Typing.Type(s,1,Array.Empty<TypeFix>());
        check("typing leading zeros",r.Table["code"] is TextColumn && r.Table["code"]!.Text(0)=="001","");
        check("typing numbers",r.Table["amount"]!.Num(0)==1250 && r.Table["amount"]!.Num(1)==0.25,"");
        check("typing trim and raw",r.Table["word"]!.Text(0)=="Abc" && r.Raw["word"]!.Text(0)==" Abc " && r.Table["_r"]!.Num(0)==2,"");
        var fixedResult=Typing.Type(s,1,new[]{new TypeFix {Column="word",Fix="upper"},new TypeFix {Column="code",Fix="padZeros",Width=5}});
        check("typing fixes",fixedResult.Table["word"]!.Text(0)=="ABC" && fixedResult.Table["code"]!.Text(0)=="00001","");
        foreach(var text in new[]{"25.10.2026","25/10/2026","10/25/2026","2026-10-25","20261025","25-Oct-2026"}) {
            var date=Typing.Type(Sheet(new[]{"date"},new[]{text}),1,Array.Empty<TypeFix>());
            check("typing date "+text,date.Table["date"]!.Type==ColType.Date && date.Table["date"]!.Num(0)==new DateTime(2026,10,25).ToOADate(),"");
        }
        check("ISO week 53",Periods.PeriodOf(new DateTime(2021,1,1).ToOADate(),Granularity.Week,new WeekRule())=="2020-W53","");
        var week=new WeekRule { Numbering="fiscal",FiscalYearStart="04-01" };
        check("fiscal week",Periods.PeriodOf(new DateTime(2026,3,30).ToOADate(),Granularity.Week,week)=="FY2026-W01","");
        var rule=new LockRule { WorkingDay=2, Holidays=new(){"2026-11-02"} };
        check("lock working days",Periods.LockTime("2026-10",Granularity.Month,rule)==new DateTime(2026,11,4,23,59,59),"");
        Table T(string?[] keys,string?[] values) => new("synthetic",new Column[]{TextColumn.From("id",keys),TextColumn.From("value",values)});
        var diff=Compare.Diff(T(new[]{"a","a","b"},new[]{"1","2","3"}),T(new[]{"a","a","c"},new[]{"1","4","5"}),new[]{"id"});
        check("diff counts and duplicate occurrence",diff.Added==1 && diff.Removed==1 && diff.Changed==1 && diff.Table.Rows==3,"");
        check("diff changed values",Enumerable.Range(0,diff.Table.Rows).Any(i=>diff.Table["_status"]!.Text(i)=="changed" && diff.Table["value"]!.Text(i)=="4" && diff.Table["value (before)"]!.Text(i)=="2"),"");
        bool missing=false; try {Compare.Diff(T(new[]{"x"},new[]{"1"}),T(new[]{"x"},new[]{"1"}),new[]{"missing"});} catch(ExcelApp.UserError) {missing=true;} check("diff missing key",missing,"");
        check("recognizer glob",Recognizer.Matches("SYN-report.XLSX","syn-*.xlsx") && !Recognizer.Matches("~$syn-report.xlsx","*.xlsx") && !Recognizer.Matches("syn.partial","*"),"");
        var detected=Recognizer.Recognize(Sheet(new[]{"title","",""},new[]{" ID "," Date ","value"}),new IngestRecipe {HeaderSignature=new(){"id","date","value"}});
        check("recognizer header detection",detected.HeaderRow==2 && detected.Score==1,"");
        Table Q(string?[] ids,string?[] dates,double[] values,string?[] categories) => new("quality",new Column[]{TextColumn.From("id",ids),TextColumn.From("date",dates),new NumberColumn("value",values,null),TextColumn.From("category",categories)});
        var quality=Quality.Analyze(Q(new string?[]{"a","a",null},new string?[]{"2026-10-01",null,"2026-10-03"},new[]{-1.0,100.0,1.0},new string?[]{"old","new","old"}),new[]{Q(new string?[]{"a","b","c"},new string?[]{"2026-09-01","2026-09-02","2026-09-03"},new[]{1.0,1.0,1.0},new string?[]{"old","old","old"})},new[]{"id"},"date",1,new[]{1},true,true);
        foreach(string kind in new[]{"duplicateKey","blank","badDate","outOfRange","outlier","newCategory","futureDate","twoPeriods"}) check("quality "+kind,quality.Flags.Any(f=>f.Kind==kind),"");
        var id=Guid.NewGuid().ToString("N"); var store=StoreRoot.ForTest(@"D:\ExcelAppTest\syn-ingest-"+id); var config=ConfigRoot.ForTest(@"D:\ExcelAppTest\syn-config-"+id);
        var fake=new FakeSource();var cat=new Catalog(store);var provider=new TableProvider(store,cat);using var queue=new WorkQueue();var notices=new List<string>();var events=new List<(string Name,System.Text.Json.JsonElement Value)>();int tick=0;using var engine=new Ingestion(store,config,fake,provider,cat,queue,(evt,payload)=>{events.Add((evt,System.Text.Json.JsonSerializer.SerializeToElement(payload,Json.Options)));if(evt=="data.notice") notices.Add(System.Text.Json.JsonSerializer.Serialize(payload));},()=>new DateTime(2027,1,15).AddSeconds(tick++));
        engine.SaveRecipe(new IngestRecipe {Id="pilot",Name="Pilot",Folders=new(){store.Folder},Patterns=new(){"*.xlsx"},HeaderSignature=new(){"id","date","value"},DateColumn="date",KeyColumns=new(){"id"}});
        LoadRecord Load(string name,params string[][] rows) {store.WriteAllText("input/"+name,name);string path=store.PathOf("input/"+name);fake.Sheets[path]=Sheet(new[]{new[]{"id","date","value"}}.Concat(rows).ToArray());return engine.Load(path,"pilot");}
        var workspace=new Workspace(store,cat,fake,queue,(name,payload)=>{});string project=workspace.Create("Synthetic project");store.WriteAllText("workspace.xlsx","v1");string workspacePath=store.PathOf("workspace.xlsx");fake.Sheets[workspacePath]=Sheet(new[]{"id","value"},new[]{"a","1"});workspace.Add(project,workspacePath);
        bool workspaceLoaded=SpinWait.SpinUntil(()=>cat.Snapshot().Workspace.Any(x=>x.Source.GroupId==project),TimeSpan.FromSeconds(10));var entry=cat.Snapshot().Workspace.FirstOrDefault(x=>x.Source.GroupId==project);check("workspace add",workspaceLoaded && entry!=null && provider.Get(new TableRef("table",entry.Info.Id)).Rows==1,"");
        int oldVersion=entry?.Source.VersionId ?? 0;fake.Sheets[workspacePath]=Sheet(new[]{"id","value"},new[]{"a","2"});store.WriteAllText("workspace.xlsx","version two");workspace.Refresh(project);bool refreshed=SpinWait.SpinUntil(()=>cat.Snapshot().Workspace.Any(x=>x.Source.GroupId==project && x.Source.VersionId!=oldVersion),TimeSpan.FromSeconds(10));check("workspace refresh",refreshed,"");
        var oct1=Load("oct1.xlsx",new[]{"a","2026-10-01","10"},new[]{"b","2026-10-02","20"});
        var nov1=Load("nov1.xlsx",new[]{"a","2026-11-01","30"});
        var oct2=Load("oct2.xlsx",new[]{"a","2026-10-01","11"},new[]{"b","2026-10-02","20"});
        var dec1=Load("dec1.xlsx",new[]{"a","2026-12-01","40"});
        var nov2=Load("nov2.xlsx",new[]{"a","2026-11-01","31"});
        check("two month pilot",new[]{oct1,nov1,oct2,dec1,nov2}.All(x=>x.Status=="loaded") && oct2.Changed==1 && oct2.Versions["2026-10"]==2 && nov2.Versions["2026-11"]==2,"");
        check("review ingestion event fields", events.Where(e=>e.Name=="data.ingest").All(e=>e.Value.TryGetProperty("file",out _) && e.Value.TryGetProperty("message",out _) && new[]{"noticed","recognized","compared","updated","notified","failed"}.Contains(e.Value.GetProperty("stage").GetString())), "");
        check("review loaded event envelope", events.Where(e=>e.Name=="data.loaded").All(e=>e.Value.TryGetProperty("record",out _)), "");
        check("ingest notice",notices.Count==5 && notices[2].Contains("updated"),"");
        var results=new ResultStore();var history=new History(store,results,rid=>engine.Recipes.FirstOrDefault(x=>x.Id==rid));var historyList=history.List("pilot","2026-10");
        check("history versions",historyList.Versions.Length==2 && historyList.Versions.Last().Current,"");
        var historyDiff=history.Diff("pilot","2026-10",1,2);check("history diff contract",historyDiff.Changed==1 && historyDiff.Columns.Any(c=>c.Name=="_status") && historyDiff.ChangedColumns.GetValueOrDefault("value")==1 && results.TryGet(historyDiff.ResultId,out _,out _),"");
        check("history trend",history.Trend("pilot","value","sum","versions","2026-10").Select(x=>x.Value).SequenceEqual(new[]{30.0,31.0}),"");
        store.WriteAllText("input/drift.xlsx","drift"); string driftPath=store.PathOf("input/drift.xlsx"); fake.Sheets[driftPath]=Sheet(new[]{"id","date","value","extra"},new[]{"a","2026-10-01","10","new"});
        var drift=engine.Load(driftPath,"pilot"); check("schema drift quarantine",drift.Status=="quarantined",""); check("schema drift reject",engine.Reject(drift.Id).Status=="rejected","");
        var truthVersions=new Versions(store,"pilot"); engine.LockDue(); check("automatic lock after cutoff",truthVersions.Snapshot().Periods["2026-10"].Locked,""); int locked=truthVersions.Snapshot().Periods["2026-10"].Current;
        var post=Load("post.xlsx",new[]{"a","2026-10-01","17"},new[]{"b","2026-10-02","20"});
        check("post close locked truth",post.Status=="loaded" && truthVersions.Snapshot().Periods["2026-10"].LockedVersion==locked && truthVersions.Snapshot().Periods["2026-10"].Versions.Last().PostClose,"");
        var duplicate=Load("duplicate.xlsx",new[]{"a","2026-10-01","17"},new[]{"b","2026-10-02","20"});check("content duplicate",duplicate.Status=="duplicate","");
        var partial=Load("partial.xlsx",new[]{"a","2026-10-01","12"});check("partial quarantine",partial.Status=="quarantined","");check("partial accept",engine.Accept(partial.Id,false,false).Status=="loaded","");
        var two=Load("two.xlsx",new[]{"a","2026-10-01","15"},new[]{"a","2026-11-01","32"});check("two periods quarantine",two.Status=="quarantined","");check("split accept",engine.Accept(two.Id,true,false).Status=="loaded","");
        var future=Load("future.xlsx",new[]{"a",new DateTime(2027,1,20).ToString("yyyy-MM-dd"),"3"});check("future quarantine",future.Status=="quarantined","");check("reject",engine.Reject(future.Id).Status=="rejected","");
        check("same file duplicate",engine.Load(store.PathOf("input/oct1.xlsx"),"pilot").Status=="duplicate","");
        var asof=provider.Get(new TableRef("report","pilot",AsOf:oct2.LoadedAt.AddMilliseconds(1).ToString("O")));
        check("as of after three loads",asof.Rows==3 && asof["value"]!.Num(0)==11,"");
        store.WriteAllText("input/drift2.xlsx","drift2");string drift2Path=store.PathOf("input/drift2.xlsx");fake.Sheets[drift2Path]=Sheet(new[]{"id","date","value","extra"},new[]{"a","2026-10-01","30","new"},new[]{"b","2026-10-02","20","new"});
        var drift2=engine.Load(drift2Path,"pilot");var acceptedDrift=engine.Accept(drift2.Id,false,true);check("schema drift accept",acceptedDrift.Status=="loaded" && engine.Recipes.First().HeaderSignature.Contains("extra"),"");
        var editable=engine.Recipes.First();editable.TypeFixes.Add(new TypeFix {Column="extra",Fix="upper"});engine.SaveRecipe(editable);
        store.WriteAllText("input/cleanup.xlsx","cleanup");string cleanupPath=store.PathOf("input/cleanup.xlsx");fake.Sheets[cleanupPath]=Sheet(new[]{"id","date","value","extra"},new[]{"a","2026-10-01","35","small"},new[]{"b","2026-10-02","20","small"});
        var cleaned=engine.Load(cleanupPath,"pilot");var undone=engine.UndoCleanup(cleaned.Id,"extra","upper");
        check("undo cleanup creates raw-derived version",undone.Status=="loaded" && provider.Get(new TableRef("report","pilot"))["extra"]!.Text(0)=="small","");
        var withDateFix=engine.Recipes.First();withDateFix.TypeFixes.Add(new TypeFix {Column="date",Fix="date",DateFormat="yyyy-MM-dd"});engine.SaveRecipe(withDateFix);
        store.WriteAllText("input/baddate.xlsx","baddate");string badPath=store.PathOf("input/baddate.xlsx");fake.Sheets[badPath]=Sheet(new[]{"id","date","value","extra"},new[]{"a","2026-10-01","36","ok"},new[]{"b","broken","20","ok"});
        var bad=engine.Load(badPath,"pilot");check("bad dates counted",bad.Quality?.Flags.Any(f=>f.Kind=="badDate" && f.Count==1)==true,"");
        foreach(var fixName in new[]{"trim","upper","padZeros","number","date","text"}) {
            var rcp=engine.Recipes.First(x=>x.Id=="pilot");rcp.TypeFixes.RemoveAll(f=>f.Column=="extra");rcp.TypeFixes.Add(new TypeFix {Column="extra",Fix=fixName,Width=4,DateFormat="yyyy-MM-dd"});engine.SaveRecipe(rcp);
            string value=fixName switch {"trim"=>" abc ","upper"=>"abc","padZeros"=>"7","number"=>"8","date"=>"2026-10-04",_=>"10"};string name="undo-"+fixName+".xlsx";store.WriteAllText("input/"+name,name);string pth=store.PathOf("input/"+name);fake.Sheets[pth]=Sheet(new[]{"id","date","value","extra"},new[]{"a","2026-10-01","50",value},new[]{"b","2026-10-02","51",value});
            var load=engine.Load(pth,"pilot");var undo=engine.UndoCleanup(load.Id,"extra",fixName);var current=provider.Get(new TableRef("report","pilot"));
            bool outcome=undo.Status=="loaded" && (fixName!="trim" || current["extra"]!.Text(0)==" abc ") && (fixName!="upper" || current["extra"]!.Text(0)=="abc") && (fixName!="padZeros" || current["extra"]!.Text(0)=="7");check("undo "+fixName,outcome,"");
        }
        foreach(var fixName in new[]{"numberFromText","dateFromText","leadingZerosKept","trimmed","mixedCase"}) {
            var rcp=engine.Recipes.First(x=>x.Id=="pilot");rcp.TypeFixes.RemoveAll(f=>f.Column=="extra");engine.SaveRecipe(rcp);
            string value=fixName switch {"numberFromText"=>"25","dateFromText"=>"2026-10-05","leadingZerosKept"=>"001","trimmed"=>" raw ",_=>"Mixed"};string name="auto-"+fixName+".xlsx";store.WriteAllText("input/"+name,name);string pth=store.PathOf("input/"+name);fake.Sheets[pth]=Sheet(new[]{"id","date","value","extra"},new[]{"a","2026-10-01","61",value},new[]{"b","2026-10-02","62",value});
            var load=engine.Load(pth,"pilot");var undo=engine.UndoCleanup(load.Id,"extra",fixName);var c=provider.Get(new TableRef("report","pilot"))["extra"]!;
            bool correct=undo.Status=="loaded" && (fixName is not ("numberFromText" or "dateFromText") || c.Type==ColType.Text) && (fixName!="leadingZerosKept" || c.Type==ColType.Number) && (fixName!="trimmed" || c.Text(0)==" raw ");check("undo "+fixName,correct,"");
        }
        engine.SaveRecipe(new IngestRecipe {Id="less-specific",Name="Other",Folders=new(){store.Folder},Patterns=new(){"*.xlsx"},HeaderSignature=new(){"id","date","other","absent"},DateColumn="date",KeyColumns=new(){"id"}});
        store.WriteAllText("best.xlsx","best");string bestPath=store.PathOf("best.xlsx");fake.Sheets[bestPath]=Sheet(new[]{"id","date","value","extra"},new[]{"a","2026-09-01","5","ok"});
        check("best recipe selected",engine.LoadBest(bestPath).RecipeId=="pilot","");
        store.WriteAllText("input/multiguard.xlsx","multiguard");string multiPath=store.PathOf("input/multiguard.xlsx");fake.Sheets[multiPath]=Sheet(new[]{"id","date","value","extra"},new[]{"a","2026-10-01","99","ok"},new[]{"z","2027-01-20","2","ok"});
        var multi=engine.Load(multiPath,"pilot");var multiFuture=engine.Accept(multi.Id,true,false);var multiPartial=engine.Accept(multi.Id,true,false);var multiDone=engine.Accept(multi.Id,true,false);
        check("separate safety approvals persist",multi.Status=="quarantined" && multiFuture.Status=="quarantined" && multiPartial.Status=="quarantined" && multiDone.Status=="loaded","");
        store.WriteAllText("input/many.xlsx","many");string manyPath=store.PathOf("input/many.xlsx");fake.Sheets[manyPath]=Sheet(new[]{new[]{"id","date","value","extra"}}.Concat(Enumerable.Range(0,25).Select(i=>new[]{"duplicate","2026-08-01",i.ToString(),"ok"})).ToArray());var many=engine.Load(manyPath,"pilot");
        var services=new DataServices {Store=store,Config=config,Emit=(eventName,payload)=>{}};var handlers=new Dictionary<string,Func<System.Text.Json.JsonElement,Task<object?>>>();DataCommands.Register(handlers,services);
        var required=new[]{"data.tables.list","data.columns","data.workspace.projects","data.workspace.create","data.workspace.rename","data.workspace.delete","data.workspace.add","data.workspace.remove","data.workspace.refresh","data.ingest.recipes","data.ingest.saveRecipe","data.ingest.deleteRecipe","data.ingest.preview","data.ingest.status","data.ingest.scanNow","data.ingest.loadFile","data.ingest.log","data.ingest.quarantine","data.ingest.accept","data.ingest.reject","data.ingest.rollback","data.ingest.undoCleanup","data.history.versions","data.history.diff","data.history.trend","data.history.lock","data.quality.rows","data.settings.get","data.settings.set"};
        check("lane A commands registered",required.All(handlers.ContainsKey) && services.Owned.Count>=2,"");
        var settingsResult=handlers["data.settings.get"](default).GetAwaiter().GetResult();var settingsJson=System.Text.Json.JsonSerializer.SerializeToElement(settingsResult,Json.Options);check("settings info contract",!settingsJson.GetProperty("pdfApproved").GetBoolean() && settingsJson.GetProperty("storeFolder").GetString()==store.Folder && settingsJson.TryGetProperty("configFolder",out _),"");
        var historyArgs=System.Text.Json.JsonSerializer.SerializeToElement(new {reportTypeId="pilot",period="2026-10",from=1,to=2},Json.Options);var routedDiff=handlers["data.history.diff"](historyArgs).GetAwaiter().GetResult();var routedJson=System.Text.Json.JsonSerializer.SerializeToElement(routedDiff,Json.Options);check("history diff command contract",routedJson.GetProperty("changed").GetInt32()==1 && routedJson.GetProperty("resultId").GetString()!.Length>0 && routedJson.GetProperty("columns").EnumerateArray().Any(c=>c.GetProperty("name").GetString()=="_status"),"");
        var settingArgs=System.Text.Json.JsonSerializer.SerializeToElement(new {settings=new Settings()},Json.Options);var setResult=handlers["data.settings.set"](settingArgs).GetAwaiter().GetResult();var setJson=System.Text.Json.JsonSerializer.SerializeToElement(setResult,Json.Options);check("settings set contract",!setJson.GetProperty("pdfApproved").GetBoolean() && setJson.GetProperty("settings").GetProperty("maxParallelExcel").GetInt32()==1,"");
        int futureFlag=multiDone.Quality!.Flags.FindIndex(f=>f.Kind=="futureDate");var qualityArgs=System.Text.Json.JsonSerializer.SerializeToElement(new {loadId=multiDone.Id,flag=futureFlag},Json.Options);var qualityResult=handlers["data.quality.rows"](qualityArgs).GetAwaiter().GetResult();var qualityJson=System.Text.Json.JsonSerializer.SerializeToElement(qualityResult,Json.Options);string qualityId=qualityJson.GetProperty("resultId").GetString()!;check("quality rows command",futureFlag>=0 && services.Results.TryGet(qualityId,out var flagged,out _) && flagged.Rows==1 && flagged["id"]!.Text(0)=="z","");
        int duplicateFlag=many.Quality!.Flags.FindIndex(f=>f.Kind=="duplicateKey");var manyArgs=System.Text.Json.JsonSerializer.SerializeToElement(new {loadId=many.Id,flag=duplicateFlag},Json.Options);var manyResult=handlers["data.quality.rows"](manyArgs).GetAwaiter().GetResult();var manyJson=System.Text.Json.JsonSerializer.SerializeToElement(manyResult,Json.Options);check("quality rows include all flags",duplicateFlag>=0 && many.Quality.Flags[duplicateFlag].SampleRows.Count==20 && manyJson.GetProperty("rows").GetInt32()==25,"");
        foreach(var owned in services.Owned.AsEnumerable().Reverse()) owned.Dispose();
        check("rollback",engine.Rollback(nov2.Id).Status=="rolledBack","");
        if(!offline) Live(check);
        } catch(Exception e) { Console.WriteLine(e); throw; }
    }
}
