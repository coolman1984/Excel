namespace ExcelApp.Data;
public static class Quality
{
    public static QualityCard Analyze(Table current,IReadOnlyList<Table> previous,IReadOnlyList<string> keys,string dateColumn,int badDate=0,IReadOnlyList<int>? badDateRows=null,bool future=false,bool twoPeriods=false,IReadOnlyList<int>? futureRows=null,IReadOnlyList<int>? periodRows=null,Action<IReadOnlyList<int>>? collectRows=null) {
        var card=new QualityCard();
        void Add(string kind,string column,List<int> rows,string message) {if(rows.Count>0) {card.Flags.Add(new QualityFlag(kind,column,rows.Count,message,rows.Take(20).ToArray()));collectRows?.Invoke(rows);}}
        if(keys.Count>0 && keys.All(k=>current[k]!=null)) {
            var groups=new Dictionary<string,List<int>>(StringComparer.Ordinal);
            for(int i=0;i<current.Rows;i++) {string key=System.Text.Json.JsonSerializer.Serialize(keys.Select(k=>current[k]!.IsNull(i)?null:current[k]!.Text(i)).ToArray());if(!groups.TryGetValue(key,out var rows))groups[key]=rows=new();rows.Add(i);}
            Add("duplicateKey",string.Join(", ",keys),groups.Values.Where(x=>x.Count>1).SelectMany(x=>x).OrderBy(x=>x).ToList(),"Some keys occur more than once.");
        }
        foreach(var c in current.Columns.Where(c=>!c.Name.StartsWith('_'))) {
            var old=previous.FirstOrDefault()?[c.Name];bool required=keys.Contains(c.Name,StringComparer.OrdinalIgnoreCase) || string.Equals(c.Name,dateColumn,StringComparison.OrdinalIgnoreCase) || old!=null && old.Rows>0 && Enumerable.Range(0,old.Rows).All(i=>!old.IsNull(i) && old.Text(i).Length>0);
            if(required) Add("blank",c.Name,Enumerable.Range(0,c.Rows).Where(i=>c.IsNull(i) || c.Text(i).Length==0).ToList(),"Some required values are blank.");
            var past=previous.Take(6).Select(t=>t[c.Name]).Where(x=>x!=null).Cast<Column>().ToArray();
            if(c.Type==ColType.Number && past.Length>0) {
                bool noNegative=past.All(p=>Enumerable.Range(0,p.Rows).Where(i=>!p.IsNull(i)).All(i=>p.Num(i)>=0));
                if(noNegative) Add("outOfRange",c.Name,Enumerable.Range(0,c.Rows).Where(i=>!c.IsNull(i) && c.Num(i)<0).ToList(),"Some values are negative for the first time.");
                var values=Enumerable.Range(0,past[0].Rows).Where(i=>!past[0].IsNull(i)).Select(i=>Math.Abs(past[0].Num(i))).OrderBy(x=>x).ToArray();
                if(values.Length>0) {double median=values.Length%2==0?(values[values.Length/2-1]+values[values.Length/2])/2:values[values.Length/2];var flagged=Enumerable.Range(0,c.Rows).Where(i=>!c.IsNull(i) && Math.Abs(c.Num(i))>=10*median && Math.Abs(c.Num(i))>0).ToHashSet();
                    var prev=previous[0];if(old!=null && keys.Count>0 && keys.All(k=>prev[k]!=null && current[k]!=null)) {string Key(Table t,int row)=>System.Text.Json.JsonSerializer.Serialize(keys.Select(k=>t[k]!.IsNull(row)?null:t[k]!.Text(row)).ToArray()); var index=new Dictionary<string,Queue<int>>();for(int i=0;i<prev.Rows;i++){string k=Key(prev,i);if(!index.TryGetValue(k,out var q))index[k]=q=new();q.Enqueue(i);}for(int i=0;i<current.Rows;i++) if(index.TryGetValue(Key(current,i),out var q)&&q.Count>0){int n=q.Dequeue();if(!c.IsNull(i)&&!old!.IsNull(n)){double prior=Math.Abs(old.Num(n));if(Math.Abs(c.Num(i)-old.Num(n))>=10*prior && Math.Abs(c.Num(i)-old.Num(n))>0)flagged.Add(i);}}}
                    Add("outlier",c.Name,flagged.OrderBy(x=>x).ToList(),"Some values differ greatly from previous values."); }
            }
            if(c.Type==ColType.Text && past.Length>0) {var seen=new HashSet<string>(StringComparer.Ordinal);foreach(var p in past) for(int i=0;i<p.Rows;i++)if(!p.IsNull(i))seen.Add(p.Text(i));var distinct=new HashSet<string>(StringComparer.Ordinal);for(int i=0;i<c.Rows;i++)if(!c.IsNull(i))distinct.Add(c.Text(i));if(distinct.Count<=200)Add("newCategory",c.Name,Enumerable.Range(0,c.Rows).Where(i=>!c.IsNull(i)&&!seen.Contains(c.Text(i))).ToList(),"Some categories are new.");}
        }
        if(badDate>0) {card.Flags.Add(new QualityFlag("badDate",dateColumn,badDate,"Some dates could not be read.",(badDateRows??Array.Empty<int>()).Take(20).ToArray()));collectRows?.Invoke(badDateRows??Array.Empty<int>());}
        if(future) {card.Flags.Add(new QualityFlag("futureDate",dateColumn,futureRows?.Count ?? 1,"Future dates were accepted.",(futureRows ?? Array.Empty<int>()).Take(20).ToArray()));collectRows?.Invoke(futureRows??Array.Empty<int>());}
        if(twoPeriods) {card.Flags.Add(new QualityFlag("twoPeriods",dateColumn,periodRows?.Count ?? 1,"Multiple periods were accepted.",(periodRows ?? Array.Empty<int>()).Take(20).ToArray()));collectRows?.Invoke(periodRows??Array.Empty<int>());}
        return card;
    }
}
