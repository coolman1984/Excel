using System.Diagnostics;
namespace ExcelApp.Data;
public sealed record HistoryDiff(string ResultId,int Rows,IReadOnlyList<ColumnInfo> Columns,double Ms,int Added,int Removed,int Changed,IReadOnlyDictionary<string,int> ChangedColumns);
public sealed record HistoryPoint(string Label,double Value,int? Version,DateTime? LoadedAt);
public sealed class History
{
    readonly StoreRoot root; readonly ResultStore results; readonly Func<string,IngestRecipe?> recipe;
    public History(StoreRoot root,ResultStore results,Func<string,IngestRecipe?> recipe) {this.root=root;this.results=results;this.recipe=recipe;}
    public (string[] Periods,VersionInfo[] Versions) List(string report,string? period=null) {
        var state=new Versions(root,report).Snapshot();var periods=state.Periods.Keys.OrderBy(x=>x,StringComparer.Ordinal).ToArray();
        var versions=state.Periods.Where(x=>period==null || x.Key==period).OrderBy(x=>x.Key,StringComparer.Ordinal).SelectMany(x=>x.Value.Versions.Select(v=>{v.Locked=x.Value.Locked && v.Version==x.Value.LockedVersion;v.Current=v.Version==x.Value.Current;return v;})).ToArray();return (periods,versions);
    }
    public HistoryDiff Diff(string report,string period,int from,int to) {
        var watch=Stopwatch.StartNew();var r=recipe(report) ?? throw new UserError("The report recipe was not found.");
        var old=TableFile.Open(root,$"reports/{report}/{period}/v{from}.xvt").Load();var next=TableFile.Open(root,$"reports/{report}/{period}/v{to}.xvt").Load();var diff=Compare.Diff(old,next,r.KeyColumns);
        var counts=new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase);foreach(var c in diff.Table.Columns.Where(x=>x.Name.EndsWith(" (before)",StringComparison.Ordinal))) {string name=c.Name[..^9];var after=diff.Table[name];int count=0;for(int i=0;i<diff.Table.Rows;i++) if(diff.Table["_status"]!.Text(i)=="changed" && after!=null && (c.IsNull(i)!=after.IsNull(i) || c.Text(i)!=after.Text(i)))count++;if(count>0)counts[name]=count;}
        string id=results.Add(diff.Table,new ResultMeta("diff",null,false,new TableRef("report",report)));watch.Stop();return new HistoryDiff(id,diff.Table.Rows,diff.Table.Columns.Where(c=>!c.Name.StartsWith('_') || c.Name=="_status").Select(c=>new ColumnInfo(c.Name,c.Type)).ToArray(),watch.Elapsed.TotalMilliseconds,diff.Added,diff.Removed,diff.Changed,counts);
    }
    public IReadOnlyList<HistoryPoint> Trend(string report,string column,string agg,string mode,string? period=null,int? last=null) {
        var state=new Versions(root,report).Snapshot();var points=new List<HistoryPoint>();
        foreach(var (p,versions) in state.Periods.OrderBy(x=>x.Key,StringComparer.Ordinal)) {if(period!=null && p!=period)continue;var chosen=mode=="versions"?versions.Versions.Where(v=>v.Status!="rolledBack").OrderBy(v=>v.Version).ToArray():versions.Versions.Where(v=>v.Version==versions.Current).ToArray();
            foreach(var v in chosen) {var t=TableFile.Open(root,$"reports/{report}/{p}/v{v.Version}.xvt").Load();var c=t[column] ?? throw new UserError("The column was not found.");double sum=0;int count=0;for(int i=0;i<t.Rows;i++)if(!c.IsNull(i)){count++;if(c.Type is ColType.Number or ColType.Date)sum+=c.Num(i);} double value=agg switch {"count"=>count,"sum"=>sum,"avg"=>count==0?0:sum/count,_=>throw new UserError("Unknown aggregate.")};points.Add(new HistoryPoint(mode=="versions"?$"{p} v{v.Version}":p,value,v.Version,v.LoadedAt));}
        }return last.HasValue?points.TakeLast(Math.Max(0,last.Value)).ToArray():points;
    }
}
