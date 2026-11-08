using System.Collections.Concurrent;
using System.Text.Json;
namespace ExcelApp.Data;
public sealed class VersionInfo
{
    public string ReportTypeId {get;set;}=""; public string Period {get;set;}=""; public int Version {get;set;} public int VersionId {get;set;}
    public string LoadId {get;set;}=""; public DateTime LoadedAt {get;set;} public int Rows {get;set;} public int Added {get;set;} public int Removed {get;set;} public int Changed {get;set;}
    public bool Current {get;set;} public bool Locked {get;set;} public bool PostClose {get;set;} public string Note {get;set;}=""; public string Status {get;set;}="current"; public SourceInfo? Source {get;set;}
}
public sealed class PeriodVersions { public List<VersionInfo> Versions {get;set;}=new(); public int Current {get;set;} public bool Locked {get;set;} public int? LockedVersion {get;set;} }
public sealed class VersionsFile { public int NextVersion {get;set;}=1; public Dictionary<string,PeriodVersions> Periods {get;set;}=new(); }
public sealed class Versions
{
    readonly StoreRoot root; readonly string relative; readonly object gate;
    static readonly ConcurrentDictionary<string,object> Gates=new(StringComparer.OrdinalIgnoreCase);
    public Versions(StoreRoot root,string report) { this.root=root; relative=$"reports/{report}/versions.json"; gate=Gates.GetOrAdd(root.PathOf(relative),_=>new object()); }
    VersionsFile Read() => root.Exists(relative) ? JsonSerializer.Deserialize<VersionsFile>(root.ReadAllText(relative),Json.Options) ?? throw new InvalidDataException("Invalid versions file.") : new();
    public VersionsFile Snapshot() {lock(gate) return Read();}
    public T Update<T>(Func<VersionsFile,T> action) {lock(gate) {using var lease=Catalog.Acquire(root,relative); var state=Read(); var result=action(state); root.WriteAllText(relative,JsonSerializer.Serialize(state,Json.Options)); return result;} }
    public int Add(string period,VersionInfo version) => Update(state=> {
        if(!state.Periods.TryGetValue(period,out var p)) state.Periods[period]=p=new();
        int n=p.Versions.Count==0?1:p.Versions.Max(v=>v.Version)+1;
        foreach(var old in p.Versions.Where(v=>v.Current)) {old.Current=false; old.Status="archived";}
        version.Period=period; version.Version=n; version.Current=true; version.Status="current"; version.PostClose=p.Locked; p.Versions.Add(version); p.Current=n; state.NextVersion=Math.Max(state.NextVersion,n+1); return n;
    });
    public void Lock(string period,bool locked) => Update(state=> {var p=state.Periods[period]; p.Locked=locked; p.LockedVersion=locked?p.Current:null; return 0;});
    public void Rollback(string period,int version) => Update(state=> {var p=state.Periods[period]; var v=p.Versions.Single(x=>x.Version==version); v.Current=false; v.Status="rolledBack"; if(p.Current==version) {var prev=p.Versions.Where(x=>x.Version<version && x.Status!="rolledBack").OrderByDescending(x=>x.Version).FirstOrDefault(); p.Current=prev?.Version ?? 0; if(prev!=null) {prev.Current=true;prev.Status="current";}} return 0;});
    public static VersionInfo? Select(PeriodVersions p,string? truth,DateTime? asOf,int? version) {
        if(version.HasValue) return p.Versions.FirstOrDefault(v=>v.Version==version);
        if(asOf.HasValue) return p.Versions.Where(v=>v.Status!="rolledBack" && v.LoadedAt<=asOf).OrderByDescending(v=>v.Version).FirstOrDefault();
        int n=truth=="locked" && p.Locked ? p.LockedVersion ?? p.Current : p.Current; return p.Versions.FirstOrDefault(v=>v.Version==n);
    }
}
