using System.Globalization;
namespace ExcelApp.Data;
public sealed class TableProvider : ITableProvider
{
    readonly StoreRoot root; readonly Catalog catalog; readonly object gate=new(); readonly Dictionary<string,Table> cached=new();
    public TableProvider(StoreRoot root,Catalog catalog) {this.root=root;this.catalog=catalog;}
    public void Invalidate(string report) {lock(gate) foreach(var key in cached.Keys.Where(k=>k.StartsWith(report+"|",StringComparison.Ordinal)).ToArray()) cached.Remove(key);}
    public Table Get(TableRef reference) {
        if(reference.Kind=="table") { var snapshot=catalog.Snapshot(); var entry=snapshot.Workspace.FirstOrDefault(x=>x.Info.Id==reference.Id); if(entry!=null) return TableFile.Open(root,entry.RelativePath).Load(); var info=snapshot.Tables.FirstOrDefault(x=>x.Kind=="table" && x.Id==reference.Id) ?? throw new UserError("The table was not found."); return TableFile.Open(root,$"tables/{info.Group}/{info.Id}.xvt").Load(); }
        var cat=catalog.Snapshot();
        if(reference.Kind=="version") {foreach(var report in cat.ReportTypes) foreach(var (period,p) in new Versions(root,report).Snapshot().Periods) {var v=p.Versions.FirstOrDefault(x=>x.VersionId.ToString(CultureInfo.InvariantCulture)==reference.Id); if(v!=null) return TableFile.Open(root,$"reports/{report}/{period}/v{v.Version}.xvt").Load();} throw new UserError("The version was not found.");}
        if(reference.Kind!="report") throw new UserError("The table was not found.");
        string key=$"{reference.Id}|{reference.Truth}|{reference.AsOf}|{reference.Version}"; lock(gate) if(cached.TryGetValue(key,out var hit)) return hit;
        DateTime? time=reference.AsOf==null?null:DateTime.Parse(reference.AsOf,CultureInfo.InvariantCulture);
        var state=new Versions(root,reference.Id).Snapshot(); var tables=new List<Table>();
        foreach(var (period,p) in state.Periods.OrderBy(x=>x.Key,StringComparer.Ordinal)) {var v=Versions.Select(p,reference.Truth,time,reference.Version); if(v!=null) tables.Add(TableFile.Open(root,$"reports/{reference.Id}/{period}/v{v.Version}.xvt").Load());}
        var union=Union(reference.Id,tables); lock(gate) cached[key]=union; return union;
    }
    public static Table Union(string name,IReadOnlyList<Table> tables) {
        if(tables.Count==0) return new Table(name,Array.Empty<Column>());
        int rows=checked(tables.Sum(x=>x.Rows)); var names=tables.SelectMany(x=>x.Columns).Select(x=>x.Name).Distinct(StringComparer.OrdinalIgnoreCase).ToArray(); var columns=new List<Column>();
        foreach(var colName in names) {var type=tables.Select(t=>t[colName]).First(c=>c!=null)!.Type;
            if(type==ColType.Text) {var vals=new string?[rows]; int offset=0; foreach(var t in tables) {var c=t[colName]; for(int i=0;i<t.Rows;i++) if(c!=null && !c.IsNull(i)) vals[offset+i]=c.Text(i); offset+=t.Rows;} columns.Add(TextColumn.From(colName,vals));}
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
