using System.Globalization;
namespace ExcelApp.Data;
public sealed class TableProvider : ITableProvider
{
    readonly StoreRoot root; readonly Catalog catalog; readonly object gate=new(); readonly Dictionary<string,Table> cached=new();
    public TableProvider(StoreRoot root,Catalog catalog) {this.root=root;this.catalog=catalog;}
    readonly List<string> order=new();   // the cached unions, oldest first
    /// <summary>How many report unions are kept in memory: a 12-period union of a wide report is large, so only the last few stay.</summary>
    public const int MaxCached=4;
    public int CachedCount {get {lock(gate) return cached.Count;}}
    public void Invalidate(string report) {lock(gate) foreach(var key in cached.Keys.Where(k=>k.StartsWith(report+"|",StringComparison.Ordinal)).ToArray()) {cached.Remove(key);order.Remove(key);}}
    public Table Get(TableRef reference) => Get(reference,null);
    /// <summary>Loads only the asked columns (plus the lineage columns _v and _r) from every period file; null loads them all.</summary>
    public Table Get(TableRef reference,IReadOnlyCollection<string>? columns) {
        string[]? want=columns==null?null:columns.Concat(new[]{"_v","_r"}).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        Table Open(string relative) {var lazy=TableFile.Open(root,relative);return want==null?lazy.Load():lazy.Load(want);}
        if(reference.Kind=="table") { var snapshot=catalog.Snapshot(); var entry=snapshot.Workspace.FirstOrDefault(x=>x.Info.Id==reference.Id); if(entry!=null) return Open(entry.RelativePath); var info=snapshot.Tables.FirstOrDefault(x=>x.Kind=="table" && x.Id==reference.Id) ?? throw new UserError("The table was not found."); return Open($"tables/{info.Group}/{info.Id}.xvt"); }
        var cat=catalog.Snapshot();
        if(reference.Kind=="version") {foreach(var report in cat.ReportTypes) foreach(var (period,p) in new Versions(root,report).Snapshot().Periods) {var v=p.Versions.FirstOrDefault(x=>x.VersionId.ToString(CultureInfo.InvariantCulture)==reference.Id); if(v!=null) return Open($"reports/{report}/{period}/v{v.Version}.xvt");} throw new UserError("The version was not found.");}
        if(reference.Kind!="report") throw new UserError("The table was not found.");
        // The key holds the period files the report points to now, so a version loaded by another process (the app, while an AI
        // agent reads) is never answered from an old union.
        var files=ReportFiles(reference);
        string baseKey=$"{reference.Id}|{reference.Truth}|{reference.AsOf}|{reference.Version}|{string.Join(";",files)}";
        string key=baseKey+"|"+(want==null?"*":string.Join(",",want.Select(x=>x.ToUpperInvariant()).OrderBy(x=>x,StringComparer.Ordinal)));
        lock(gate) {if(cached.TryGetValue(key,out var hit)) return hit; if(cached.TryGetValue(baseKey+"|*",out var full)) return full;}
        DateTime? time=reference.AsOf==null?null:DateTime.Parse(reference.AsOf,CultureInfo.InvariantCulture);
        var tables=files.Select(Open).ToList();
        var union=Union(reference.Id,tables);
        lock(gate) {cached[key]=union; order.Remove(key); order.Add(key); while(order.Count>MaxCached) {cached.Remove(order[0]); order.RemoveAt(0);}}
        return union;
    }
    /// <summary>The period files a report reference reads, in union order (the same order Get concatenates them).</summary>
    List<string> ReportFiles(TableRef reference) {
        DateTime? time=reference.AsOf==null?null:DateTime.Parse(reference.AsOf,CultureInfo.InvariantCulture);
        var state=new Versions(root,reference.Id).Snapshot(); var files=new List<string>();
        foreach(var (period,p) in state.Periods.OrderBy(x=>x.Key,StringComparer.Ordinal)) {var v=Versions.Select(p,reference.Truth,time,reference.Version); if(v!=null) files.Add($"reports/{reference.Id}/{period}/v{v.Version}.xvt");}
        return files;
    }
    /// <summary>The single file of a table or version reference; null for a report (several period files).</summary>
    string? SingleFile(TableRef reference) {
        if(reference.Kind=="table") { var snapshot=catalog.Snapshot(); var entry=snapshot.Workspace.FirstOrDefault(x=>x.Info.Id==reference.Id); if(entry!=null) return entry.RelativePath; var info=snapshot.Tables.FirstOrDefault(x=>x.Kind=="table" && x.Id==reference.Id) ?? throw new UserError("The table was not found."); return $"tables/{info.Group}/{info.Id}.xvt"; }
        if(reference.Kind=="version") {foreach(var report in catalog.Snapshot().ReportTypes) foreach(var (period,p) in new Versions(root,report).Snapshot().Periods) {var v=p.Versions.FirstOrDefault(x=>x.VersionId.ToString(CultureInfo.InvariantCulture)==reference.Id); if(v!=null) return $"reports/{report}/{period}/v{v.Version}.xvt";} throw new UserError("The version was not found.");}
        if(reference.Kind!="report") throw new UserError("The table was not found.");
        return null;
    }
    /// <summary>Only these rows (row numbers of Get's table, in this order) of these columns. Each period file reads its own rows.</summary>
    public Table GetRows(TableRef reference,IReadOnlyCollection<string>? columns,int[] rows) {
        string[]? want=columns==null?null:columns.Concat(new[]{"_v","_r"}).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var single=SingleFile(reference);
        if(single!=null) return TableFile.Open(root,single).LoadRows(want,rows);
        var files=ReportFiles(reference).Select(f=>TableFile.Open(root,f)).ToList();
        var starts=new int[files.Count+1]; for(int i=0;i<files.Count;i++) starts[i+1]=checked(starts[i]+files[i].Rows);
        var local=files.Select(_=>new List<int>()).ToArray(); var where=new (int File,int Index)[rows.Length];
        for(int k=0;k<rows.Length;k++) {int f=Array.BinarySearch(starts,rows[k]); if(f<0) f=~f-1; while(f+1<starts.Length-1 && starts[f+1]==rows[k] && files[f].Rows==0) f++; if(f<0||f>=files.Count||rows[k]>=starts[^1]) throw new ArgumentOutOfRangeException(nameof(rows)); where[k]=(f,local[f].Count); local[f].Add(rows[k]-starts[f]);}
        var parts=new List<Table>(); var partStart=new int[files.Count]; int at=0;
        for(int f=0;f<files.Count;f++) {partStart[f]=at; parts.Add(files[f].LoadRows(want,local[f].ToArray())); at+=local[f].Count;}
        var union=Union(reference.Id,parts);
        var order=where.Select(w=>partStart[w.File]+w.Index).ToArray();
        return new Table(union.Name,union.Columns.Select(c=>ExcelApp.Data.Query.Engine.Gather(c,order)).ToArray());
    }
    public static Table Union(string name,IReadOnlyList<Table> tables) {
        if(tables.Count==0) return new Table(name,Array.Empty<Column>());
        int rows=checked(tables.Sum(x=>x.Rows)); var names=tables.SelectMany(x=>x.Columns).Select(x=>x.Name).Distinct(StringComparer.OrdinalIgnoreCase).ToArray(); var columns=new List<Column>();
        foreach(var colName in names) {var types=tables.Select(t=>t[colName]).Where(c=>c!=null).Select(c=>c!.Type).Distinct().ToArray();var type=types.Length==1?types[0]:ColType.Text;
            if(type==ColType.Text && tables.All(t=>t[colName] is null or TextColumn)) {
                // Merge the period dictionaries instead of making one string per row: 12 million rows stay a few int arrays.
                var map=new Dictionary<string,int>(StringComparer.Ordinal); var dict=new List<string>(); var codes=new int[rows]; int offset=0;
                foreach(var t in tables) {var c=t[colName] as TextColumn;
                    if(c==null) {Array.Fill(codes,-1,offset,t.Rows); offset+=t.Rows; continue;}
                    var remap=new int[c.Dict.Length]; for(int d=0;d<remap.Length;d++) {if(!map.TryGetValue(c.Dict[d],out int id)) {id=dict.Count; map[c.Dict[d]]=id; dict.Add(c.Dict[d]);} remap[d]=id;}
                    for(int i=0;i<t.Rows;i++) {int code=c.Codes[i]; codes[offset+i]=code<0?-1:remap[code];} offset+=t.Rows;}
                columns.Add(new TextColumn(colName,codes,dict.ToArray()));}
            else if(type==ColType.Text) {var vals=new string?[rows]; int offset=0; foreach(var t in tables) {var c=t[colName]; for(int i=0;i<t.Rows;i++) if(c!=null && !c.IsNull(i)) vals[offset+i]=c.Text(i); offset+=t.Rows;} columns.Add(TextColumn.From(colName,vals));}
            else if(type==ColType.Bool) {var vals=new sbyte[rows]; Array.Fill(vals,(sbyte)-1); int offset=0; foreach(var t in tables) {var c=t[colName]; for(int i=0;i<t.Rows;i++) if(c!=null && !c.IsNull(i)) vals[offset+i]=(sbyte)c.Num(i); offset+=t.Rows;} columns.Add(new BoolColumn(colName,vals));}
            else {var vals=new double[rows]; var nulls=new bool[rows]; int offset=0; foreach(var t in tables) {var c=t[colName]; for(int i=0;i<t.Rows;i++) {nulls[offset+i]=c==null || c.IsNull(i); if(!nulls[offset+i]) vals[offset+i]=c!.Num(i);} offset+=t.Rows;} columns.Add(new NumberColumn(colName,vals,nulls.Any(x=>x)?nulls:null,type));}
        } return new Table(name,columns);
    }
    public IReadOnlyList<TableInfo> List() {
        var c=catalog.Snapshot();var result=c.Tables.Concat(c.Workspace.Select(x=>x.Info)).ToList();
        foreach(var report in c.ReportTypes) {var versions=new Versions(root,report).Snapshot();int rows=0;DateTime loaded=DateTime.MinValue;var columns=new Dictionary<string,ColumnInfo>(StringComparer.OrdinalIgnoreCase);
            foreach(var (period,p) in versions.Periods) {var v=Versions.Select(p,"latest",null,null);if(v==null)continue;var table=TableFile.Open(root,$"reports/{report}/{period}/v{v.Version}.xvt"); rows=checked(rows+table.Rows);if(v.LoadedAt>loaded)loaded=v.LoadedAt;foreach(var col in table.Columns) columns.TryAdd(col.Name,col);}
            result.Add(new TableInfo("report",report,report,report,rows,columns.Values.Where(x=>!x.Name.StartsWith('_')).ToArray(),loaded));
        }return result;
    }
    public SourceInfo? SourceOf(int versionId) {var workspace=catalog.Snapshot().Workspace.FirstOrDefault(x=>x.Source.VersionId==versionId);if(workspace!=null)return workspace.Source;foreach(var report in catalog.Snapshot().ReportTypes) foreach(var p in new Versions(root,report).Snapshot().Periods.Values) {var source=p.Versions.FirstOrDefault(v=>v.VersionId==versionId)?.Source; if(source!=null) return source;} return null;}
    public int SaveAsVersion(string reportTypeId,string period,Table data,string note) {
        var versions=new Versions(root,reportTypeId); int id=catalog.NextVersion();
        var cols=data.Columns.Where(c=>c.Name!="_v").ToList(); cols.Add(new NumberColumn("_v",Enumerable.Repeat((double)id,data.Rows).ToArray(),null));
        int actual=versions.Update(state=> {
            if(!state.Periods.TryGetValue(period,out var p)) state.Periods[period]=p=new();
            int next=p.Versions.Count==0?1:p.Versions.Max(v=>v.Version)+1;
            TableFile.Write(root,$"reports/{reportTypeId}/{period}/v{next}.xvt",new Table(data.Name,cols),new Dictionary<string,string>());
            foreach(var old in p.Versions.Where(v=>v.Current)) {old.Current=false;old.Status="archived";}
            p.Versions.Add(new VersionInfo {ReportTypeId=reportTypeId, Period=period,Version=next,VersionId=id,LoadedAt=DateTime.Now,Rows=data.Rows,Note=note,Current=true,PostClose=p.Locked}); p.Current=next; state.NextVersion=Math.Max(state.NextVersion,next+1); return next;
        }); Invalidate(reportTypeId); return actual;
    }
}
