using Timer = System.Threading.Timer;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
namespace ExcelApp.Data;
public sealed class Ingestion : IIngestion, IDisposable
{
    readonly StoreRoot root; readonly Recipes recipes; readonly IRawSource source; readonly TableProvider provider; readonly Catalog catalog; readonly WorkQueue queue; readonly Action<string,object?> emit; readonly object gate=new(); readonly Func<DateTime> now; Watcher? watcher; Timer? timer;
    public event Action<LoadRecord>? Loaded;
    public IReadOnlyList<IngestRecipe> Recipes => recipes.List();
    public Ingestion(StoreRoot root,ConfigRoot config,IRawSource source,TableProvider provider,Catalog catalog,WorkQueue queue,Action<string,object?> emit,Func<DateTime>? clock=null) {now=clock ?? (()=>DateTime.Now);this.root=root;recipes=new Recipes(config);this.source=source;this.provider=provider;this.catalog=catalog;this.queue=queue;this.emit=emit;}
    public void SaveRecipe(IngestRecipe r) {recipes.Save(r);catalog.Update(c=>{if(!c.ReportTypes.Contains(r.Id)) c.ReportTypes.Add(r.Id);return 0;});RefreshWatcher();}
    public void DeleteRecipe(string id) {recipes.Delete(id);RefreshWatcher();}
    void RefreshWatcher() {lock(gate) {if(watcher==null)return;watcher.Dispose();watcher=MakeWatcher();watcher.Start();}}
    Watcher MakeWatcher()=>new Watcher(()=>Recipes,(path,r)=>queue.Enqueue("ingestion",_=>{LoadBest(path);return Task.CompletedTask;}));
    public void Start() {lock(gate) {if(watcher!=null)return;watcher=MakeWatcher();watcher.Start();timer=new Timer(_=>LockDue(),null,TimeSpan.Zero,TimeSpan.FromHours(1));}}
    public void Stop() {lock(gate) {watcher?.Dispose();watcher=null;timer?.Dispose();timer=null;}}
    public void Dispose()=>Stop();
    public void LockDue() {foreach(var r in Recipes) foreach(var (period,p) in new Versions(root,r.Id).Snapshot().Periods) if(!p.Locked && p.Current>0 && now()>=Periods.LockTime(period,r.Granularity,r.Lock,r.Week)) new Versions(root,r.Id).Lock(period,true);}
    void Emit(string name,object? payload) {try {emit(name,payload);} catch(Exception) { }}
    void NotifyLoaded(LoadRecord record) {if(Loaded==null)return;foreach(Action<LoadRecord> handler in Loaded.GetInvocationList()) try {handler(record);} catch(Exception) { }}
    void Stage(LoadRecord record,string stage) => Emit("data.ingest",new {loadId=record.Id,stage});
    void Log(LoadRecord record) => root.AppendLine("ingest/log.jsonl",JsonSerializer.Serialize(record,new JsonSerializerOptions(Json.Options) {WriteIndented=false}));
    public IReadOnlyList<LoadRecord> LogRecords() {if(!root.Exists("ingest/log.jsonl")) return Array.Empty<LoadRecord>(); return root.ReadAllText("ingest/log.jsonl").Split('\n',StringSplitOptions.RemoveEmptyEntries).Select(x=>JsonSerializer.Deserialize<LoadRecord>(x,Json.Options)!).Where(x=>x!=null).GroupBy(x=>x.Id).Select(g=>g.Last()).Reverse().ToArray();}
    static string HashFile(string path) {using var stream=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.Read);return Convert.ToHexString(SHA256.HashData(stream));}
    static string ContentHash(Table t) {using var hash=IncrementalHash.CreateHash(HashAlgorithmName.SHA256);foreach(var col in t.Columns.Where(c=>!c.Name.StartsWith('_'))) {hash.AppendData(Encoding.UTF8.GetBytes(col.Name));for(int i=0;i<t.Rows;i++) {var v=col.IsNull(i)?"N":"V"+col.Text(i);var b=Encoding.UTF8.GetBytes(v);hash.AppendData(BitConverter.GetBytes(b.Length));hash.AppendData(b);}}return Convert.ToHexString(hash.GetHashAndReset());}
    Dictionary<string,string> Seen() => root.Exists("ingest/seen.json") ? JsonSerializer.Deserialize<Dictionary<string,string>>(root.ReadAllText("ingest/seen.json"),Json.Options) ?? new() : new();
    public void ScanNow() {lock(gate) watcher?.Scan();}
    public object Status() {var status=queue.Status;return new {running=watcher!=null,watching=Recipes.SelectMany(r=>r.Folders).Distinct(StringComparer.OrdinalIgnoreCase).Select(folder=>new {folder,ok=Directory.Exists(folder),error=Directory.Exists(folder)?null:"Folder is unavailable."}).ToArray(),queue=status.Pending,current=status.Current};}
    public void QueueLoad(string path,string recipeId) => queue.Enqueue("ingestion",_=>{Load(path,recipeId);return Task.CompletedTask;});
    public LoadRecord LoadBest(string path) {
        var candidates=Recipes.Where(r=>r.Folders.Any(f=>string.Equals(Path.GetFullPath(f).TrimEnd('\\'),Path.GetDirectoryName(path)?.TrimEnd('\\'),StringComparison.OrdinalIgnoreCase)) && r.Patterns.Any(p=>Recognizer.Matches(Path.GetFileName(path),p))).ToArray();
        if(candidates.Length==0) throw new UserError("No recipe matches this file.");
        var sheets=new Dictionary<string,RawSheet>(StringComparer.OrdinalIgnoreCase); IngestRecipe? best=null; double score=-1;
        foreach(var recipe in candidates) {try {string sheet=recipe.Sheet ?? "";if(!sheets.TryGetValue(sheet,out var raw)) sheets[sheet]=raw=source.ReadSheet(path,sheet,CancellationToken.None);double found=Recognizer.Recognize(raw,recipe).Score;if(found>score) {score=found;best=recipe;}} catch(Exception) { }}
        return Load(path,(best ?? candidates[0]).Id);
    }
    public LoadRecord Load(string path,string recipeId) {
        var r=Recipes.First(x=>x.Id==recipeId);var rec=new LoadRecord {Id="L"+Guid.NewGuid().ToString("N"),RecipeId=recipeId,SourcePath=path,LoadedAt=now()};
        lock(gate) {try {Stage(rec,"noticed");rec.FileHash=HashFile(path);var seen=Seen(); if(seen.ContainsKey(recipeId+":"+rec.FileHash)) {rec.Status="duplicate";Log(rec);return rec;}
            var rawSheet=source.ReadSheet(path,r.Sheet,CancellationToken.None);var recognition=Recognizer.Recognize(rawSheet,r); var typed=Typing.Type(rawSheet,recognition.HeaderRow,r.TypeFixes);rec.Cleanup=typed.Cleanup;rec.Rows=typed.Table.Rows;rec.ContentHash=ContentHash(typed.Table);
            var map=new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase);for(int i=0;i<rawSheet.Cols;i++) map[rawSheet.Text(recognition.HeaderRow-rawSheet.FirstRow,i).Trim()]=i+rawSheet.FirstCol;
            Stage(rec,"recognised");
            if(recognition.Score<0.8 || recognition.Missing.Count>0 || recognition.Added.Count>0) return Quarantine(rec,typed.Table,typed.Raw,recognition.Score<0.8?"The columns do not match the recipe":"The columns have changed.",rawSheet.Name,recognition.HeaderRow,map);
            return Process(rec,r,typed.Table,typed.Raw,rawSheet.Name,recognition.HeaderRow,map,false,false);
        } catch(Exception e) when(e is not OperationCanceledException) {rec.Status="failed";rec.Reason="The file could not be loaded.";Log(rec);return rec;} }
    }
    sealed class StageInfo {public LoadRecord Record {get;set;}=new();public string Sheet {get;set;}="";public int Header {get;set;} public Dictionary<string,int> Map {get;set;}=new();public List<string> Approved {get;set;}=new();}
    LoadRecord Quarantine(LoadRecord rec,Table data,Table raw,string reason,string sheet,int header,Dictionary<string,int> map,HashSet<string>? approved=null) {rec.Status="quarantined";rec.Reason=reason;TableFile.Write(root,$"ingest/quarantine/{rec.Id}.xvt",data,new Dictionary<string,string>());TableFile.Write(root,$"ingest/quarantine/{rec.Id}.raw.xvt",raw,new Dictionary<string,string>());root.WriteAllText($"ingest/quarantine/{rec.Id}.json",JsonSerializer.Serialize(new StageInfo {Record=rec,Sheet=sheet,Header=header,Map=map,Approved=approved?.ToList() ?? new()},Json.Options));Log(rec);Emit("data.quarantine",rec);return rec;}
    public static Table Slice(Table table,IReadOnlyList<int> rows) {var cols=new List<Column>();foreach(var c in table.Columns) {if(c.Type==ColType.Text) cols.Add(TextColumn.From(c.Name,rows.Select(i=>c.IsNull(i)?null:c.Text(i)).ToArray()));else if(c.Type==ColType.Bool) cols.Add(new BoolColumn(c.Name,rows.Select(i=>c.IsNull(i)?(sbyte)-1:(sbyte)c.Num(i)).ToArray()));else {var vals=rows.Select(i=>c.IsNull(i)?0:c.Num(i)).ToArray();var nulls=rows.Select(c.IsNull).ToArray();cols.Add(new NumberColumn(c.Name,vals,nulls.Any(x=>x)?nulls:null,c.Type));}}return new Table(table.Name,cols);}
    LoadRecord Process(LoadRecord rec,IngestRecipe recipe,Table data,Table raw,string sheet,int header,Dictionary<string,int> map,bool accepted,bool split,HashSet<string>? approved=null,bool forceVersion=false) {
        approved ??= new HashSet<string>(StringComparer.Ordinal);rec.Quality=null;
        var date=data[recipe.DateColumn];if(date==null) return Quarantine(rec,data,raw,"The date column is missing.",sheet,header,map,approved);
        var groups=new SortedDictionary<string,List<int>>(StringComparer.Ordinal);bool future=false;int badDate=0;var badRows=new List<int>();
        for(int i=0;i<data.Rows;i++) {if(date.IsNull(i) || double.IsNaN(date.Num(i))) {badDate++;badRows.Add(i);continue;} DateTime d;try {d=DateTime.FromOADate(date.Num(i));} catch(ArgumentException) {badDate++;badRows.Add(i);continue;} if(d.Date>now().Date.AddDays(1)) future=true; string p=Periods.PeriodOf(date.Num(i),recipe.Granularity,recipe.Week);if(!groups.TryGetValue(p,out var rows)) groups[p]=rows=new();rows.Add(i);}
        if(badDate>0) {rec.Quality ??= new QualityCard();rec.Quality.Flags.Add(new QualityFlag("badDate",recipe.DateColumn,badDate,"Some dates could not be read.",badRows.Take(20).ToArray()));}
        if(groups.Count==0) return Quarantine(rec,data,raw,"No readable dates were found.",sheet,header,map,approved);
        if(groups.Count>1 && (!accepted || !split)) return Quarantine(rec,data,raw,"This file contains more than one period.",sheet,header,map,approved);
        if(future && !approved.Contains("future")) return Quarantine(rec,data,raw,"This file contains future dates.",sheet,header,map,approved);
        var versions=new Versions(root,recipe.Id);var state=versions.Snapshot();
        foreach(var (period,rows) in groups) {if(!state.Periods.TryGetValue(period,out var p) || p.Current==0) continue; var old=TableFile.Open(root,$"reports/{recipe.Id}/{period}/v{p.Current}.xvt").Load();if(!approved.Contains("partial") && rows.Count<old.Rows*(1-recipe.PartialThreshold)) return Quarantine(rec,data,raw,$"This file has far fewer rows than the version it would replace ({rows.Count} against {old.Rows}).",sheet,header,map,approved);}
        if(!forceVersion && groups.All(entry => state.Periods.TryGetValue(entry.Key,out var pv) && pv.Current>0 && ContentHash(Slice(data,entry.Value))==ContentHash(TableFile.Open(root,$"reports/{recipe.Id}/{entry.Key}/v{pv.Current}.xvt").Load()))) {rec.Status="duplicate"; Log(rec); return rec;}
        Stage(rec,"compared"); var notices=new List<string>();var qualityRows=new Dictionary<int,int[]>();if(badDate>0) qualityRows[0]=badRows.ToArray();
        foreach(var (period,rows) in groups) {var part=Slice(data,rows);var rawPart=Slice(raw,rows);bool existed=state.Periods.TryGetValue(period,out var oldState) && oldState.Current>0;Table old=existed?TableFile.Open(root,$"reports/{recipe.Id}/{period}/v{oldState!.Current}.xvt").Load():new Table(part.Name,part.Columns.Select(c=>c.Type==ColType.Text?(Column)TextColumn.From(c.Name,Array.Empty<string?>()):c.Type==ColType.Bool?new BoolColumn(c.Name,Array.Empty<sbyte>()):new NumberColumn(c.Name,Array.Empty<double>(),null,c.Type)).ToArray());
            var history=new List<Table>(); if(oldState!=null) foreach(var historical in oldState.Versions.Where(x=>x.Status!="rolledBack").OrderByDescending(x=>x.Version).Take(6)) history.Add(TableFile.Open(root,$"reports/{recipe.Id}/{period}/v{historical.Version}.xvt").Load());
            var futureRows=accepted && future ? Enumerable.Range(0,rows.Count).Where(i=>DateTime.FromOADate(date.Num(rows[i])).Date>now().Date.AddDays(1)).ToArray() : Array.Empty<int>();var periodRows=accepted && groups.Count>1?Enumerable.Range(0,rows.Count).ToArray():Array.Empty<int>();
            var captured=new List<IReadOnlyList<int>>();var quality=Quality.Analyze(part,history,recipe.KeyColumns,recipe.DateColumn,0,null,futureRows.Length>0,periodRows.Length>0,futureRows,periodRows,x=>captured.Add(x)); rec.Quality ??= new QualityCard();for(int qi=0;qi<quality.Flags.Count;qi++) {int index=rec.Quality.Flags.Count;var flag=quality.Flags[qi];qualityRows[index]=captured[qi].Select(i=>rows[i]).ToArray();rec.Quality.Flags.Add(flag with {SampleRows=flag.SampleRows.Select(i=>rows[i]).ToArray()});}
            var diff=Compare.Diff(old,part,recipe.KeyColumns);rec.Added+=diff.Added;rec.Removed+=diff.Removed;rec.Changed+=diff.Changed;
            int id=catalog.NextVersion();int n=versions.Update(vs=> {if(!vs.Periods.TryGetValue(period,out var pv))vs.Periods[period]=pv=new();int next=pv.Versions.Count==0?1:pv.Versions.Max(x=>x.Version)+1;var cols=part.Columns.Where(c=>c.Name!="_v").ToList();cols.Add(new NumberColumn("_v",Enumerable.Repeat((double)id,part.Rows).ToArray(),null));TableFile.Write(root,$"reports/{recipe.Id}/{period}/v{next}.xvt",new Table(part.Name,cols),new Dictionary<string,string>());TableFile.Write(root,$"reports/{recipe.Id}/{period}/v{next}.raw.xvt",rawPart,new Dictionary<string,string>());TableFile.Write(root,$"reports/{recipe.Id}/{period}/v{next}.diff.xvt",diff.Table,new Dictionary<string,string>());foreach(var prev in pv.Versions.Where(x=>x.Current)){prev.Current=false;prev.Status="archived";}pv.Versions.Add(new VersionInfo {ReportTypeId=recipe.Id,Period=period,Version=next,VersionId=id,LoadId=rec.Id,LoadedAt=rec.LoadedAt,Rows=part.Rows,Added=diff.Added,Removed=diff.Removed,Changed=diff.Changed,Current=true,PostClose=pv.Locked,Source=new SourceInfo(id,"report",recipe.Id,period,next,rec.SourcePath,sheet,header,rec.LoadedAt,rec.Id,map)});pv.Current=next;return next;});rec.Periods.Add(period);rec.Versions[period]=n;string label=DateTime.TryParseExact(period+"-01","yyyy-MM-dd",System.Globalization.CultureInfo.InvariantCulture,System.Globalization.DateTimeStyles.None,out var month)?month.ToString("MMMM",System.Globalization.CultureInfo.InvariantCulture):period;notices.Add(label+(existed?$" updated ({diff.Changed} changed, {diff.Added} added)":" added"));}
        TableFile.Write(root,$"ingest/quality/{rec.Id}.xvt",data,new Dictionary<string,string>());foreach(var (index,indices) in qualityRows) root.WriteAllText($"ingest/quality/{rec.Id}.{index}.json",JsonSerializer.Serialize(indices,Json.Options));rec.Status="loaded";rec.Reason=null;var seen=Seen();seen[recipe.Id+":"+rec.FileHash]=rec.Id;root.WriteAllText("ingest/seen.json",JsonSerializer.Serialize(seen,Json.Options));provider.Invalidate(recipe.Id);Stage(rec,"updated");Log(rec);NotifyLoaded(rec);Emit("data.loaded",rec);Emit("data.notice",new {message=string.Join(", ",notices)+"."});Stage(rec,"notified");return rec;
    }
    public LoadRecord Accept(string id,bool split,bool updateRecipe) {lock(gate) {var info=JsonSerializer.Deserialize<StageInfo>(root.ReadAllText($"ingest/quarantine/{id}.json"),Json.Options)!;var recipe=Recipes.First(x=>x.Id==info.Record.RecipeId);if((info.Record.Reason=="The columns have changed." || info.Record.Reason=="The columns do not match the recipe") && !updateRecipe) throw new UserError("Update the recipe to accept the changed columns.");var data=TableFile.Open(root,$"ingest/quarantine/{id}.xvt").Load();var raw=TableFile.Open(root,$"ingest/quarantine/{id}.raw.xvt").Load();if(updateRecipe) {recipe.HeaderSignature=data.Columns.Where(x=>!x.Name.StartsWith('_')).Select(x=>x.Name).ToList();SaveRecipe(recipe);}var approved=new HashSet<string>(info.Approved,StringComparer.Ordinal);if(info.Record.Reason=="This file contains future dates.") approved.Add("future");if(info.Record.Reason?.StartsWith("This file has far fewer rows",StringComparison.Ordinal)==true) approved.Add("partial");var rec=Process(info.Record,recipe,data,raw,info.Sheet,info.Header,info.Map,true,split,approved);if(rec.Status=="loaded") DeleteStage(id);return rec;}}
    public LoadRecord UndoCleanup(string loadId,string column,string fix) {
        lock(gate) {
            var previous=LogRecords().First(x=>x.Id==loadId); var recipe=Recipes.First(x=>x.Id==previous.RecipeId);
            recipe.TypeFixes.RemoveAll(x=>string.Equals(x.Column,column,StringComparison.OrdinalIgnoreCase) && string.Equals(x.Fix,fix,StringComparison.OrdinalIgnoreCase));
            if(fix is "numberFromText" or "dateFromText") recipe.TypeFixes.Add(new TypeFix {Column=column,Fix="text"});
            SaveRecipe(recipe); LoadRecord? last=null;
            foreach(var (period,number) in previous.Versions) {
                var raw=TableFile.Open(root,$"reports/{recipe.Id}/{period}/v{number}.raw.xvt").Load();var columns=raw.Columns.Where(c=>!c.Name.StartsWith('_')).ToArray(); int count=raw.Rows;
                var kinds=new byte[count+1,columns.Length];var nums=new double[count+1,columns.Length];var texts=new string[count+1,columns.Length];var dates=new bool[count+1,columns.Length];
                for(int c=0;c<columns.Length;c++) {texts[0,c]=columns[c].Name;kinds[0,c]=2;for(int r=0;r<count;r++) {texts[r+1,c]=columns[c].IsNull(r)?"":columns[c].Text(r);kinds[r+1,c]=(byte)(texts[r+1,c].Length==0?0:2);}}
                var typed=Typing.Type(new RawSheet(kinds,nums,texts,dates) {Name=raw.Name},1,recipe.TypeFixes,column,fix);
                var revisedColumns=typed.Table.Columns.Select(c=>c.Name=="_r" && raw["_r"]!=null ? raw["_r"]! : c).ToArray();
                var record=new LoadRecord {Id="L"+Guid.NewGuid().ToString("N"),RecipeId=recipe.Id,SourcePath=previous.SourcePath,FileHash=previous.FileHash,LoadedAt=now(),Rows=count,Cleanup=typed.Cleanup};
                var originalSource=new Versions(root,recipe.Id).Snapshot().Periods[period].Versions.First(x=>x.Version==number).Source;
                last=Process(record,recipe,new Table(raw.Name,revisedColumns),raw,originalSource?.Sheet ?? "",originalSource?.HeaderRow ?? 1,originalSource==null?new Dictionary<string,int>():new Dictionary<string,int>(originalSource.SourceColumn),true,true,forceVersion:true);
            }
            return last ?? throw new UserError("The load has no versions to restore.");
        }
    }
    void DeleteStage(string id) {root.Delete($"ingest/quarantine/{id}.json");root.Delete($"ingest/quarantine/{id}.xvt");root.Delete($"ingest/quarantine/{id}.raw.xvt");}
    public LoadRecord Reject(string id) {lock(gate) {var info=JsonSerializer.Deserialize<StageInfo>(root.ReadAllText($"ingest/quarantine/{id}.json"),Json.Options)!;DeleteStage(id);info.Record.Status="rejected";Log(info.Record);return info.Record;}}
    public LoadRecord Rollback(string id) {lock(gate) {var rec=LogRecords().First(x=>x.Id==id);foreach(var (period,version) in rec.Versions) new Versions(root,rec.RecipeId).Rollback(period,version);rec.Status="rolledBack";provider.Invalidate(rec.RecipeId);Log(rec);NotifyLoaded(rec);Emit("data.loaded",rec);return rec;}}
}
