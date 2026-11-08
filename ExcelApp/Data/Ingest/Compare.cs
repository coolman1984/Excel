using System.Text.Json;
namespace ExcelApp.Data;
public sealed record DiffOutcome(int Added,int Removed,int Changed,Table Table);
public static class Compare
{
    public static DiffOutcome Diff(Table oldT,Table newT,IReadOnlyList<string> keys) {
        if(keys.Count==0 || keys.Any(k=>oldT[k]==null || newT[k]==null)) throw new UserError("A key column is missing.");
        var names=oldT.Columns.Concat(newT.Columns).Select(c=>c.Name).Where(n=>!n.StartsWith('_') && !keys.Contains(n,StringComparer.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        string Key(Table t,int row) => JsonSerializer.Serialize(keys.Select(k=>t[k]!.IsNull(row)?null:t[k]!.Text(row)).ToArray());
        var oldRows=new Dictionary<string,Queue<int>>(StringComparer.Ordinal);
        for(int i=0;i<oldT.Rows;i++) {var key=Key(oldT,i); if(!oldRows.TryGetValue(key,out var queue)) oldRows[key]=queue=new(); queue.Enqueue(i);}
        var matched=new bool[oldT.Rows]; var output=new List<(int? Old,int? New,string Status)>(); int added=0,removed=0,changed=0;
        for(int n=0;n<newT.Rows;n++) {var key=Key(newT,n); if(!oldRows.TryGetValue(key,out var queue) || queue.Count==0) {output.Add((null,n,"added"));added++;continue;} int o=queue.Dequeue();matched[o]=true;
            bool different=names.Any(name=>Value(oldT,name,o)!=Value(newT,name,n)); if(different) {changed++;output.Add((o,n,"changed"));}
        }
        for(int o=0;o<oldT.Rows;o++) if(!matched[o]) {removed++;output.Add((o,null,"removed"));}
        var cols=new List<Column>();
        foreach(var key in keys) cols.Add(TextColumn.From(key,output.Select(x=>x.New.HasValue?Value(newT,key,x.New.Value):Value(oldT,key,x.Old!.Value)).ToArray()));
        cols.Add(TextColumn.From("_status",output.Select(x=>(string?)x.Status).ToArray()));
        foreach(var name in names) {cols.Add(TextColumn.From(name,output.Select(x=>x.New.HasValue?Value(newT,name,x.New.Value):null).ToArray())); cols.Add(TextColumn.From(name+" (before)",output.Select(x=>x.Old.HasValue?Value(oldT,name,x.Old.Value):null).ToArray()));}
        return new DiffOutcome(added,removed,changed,new Table("diff",cols));
    }
    static string? Value(Table t,string col,int row) {var c=t[col];return c==null || c.IsNull(row)?null:c.Text(row);}
}
