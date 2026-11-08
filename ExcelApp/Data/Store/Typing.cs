using System.Globalization;
using System.Text.RegularExpressions;
using ExcelViewer.Core;
namespace ExcelApp.Data;
public sealed record TypedResult(Table Table, CleanupReport Cleanup, Table Raw);
public static class Typing
{
    static readonly string[] Formats = { "dd.MM.yyyy", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "yyyyMMdd", "dd-MMM-yyyy" };
    static bool Numeric(string s, out double number) {
        s=s.Trim(); bool percent=s.EndsWith('%'); if(percent) s=s[..^1];
        return Parse(s,percent,out number);
    }
    static bool Parse(string s,bool percent,out double n) {
        const NumberStyles style=NumberStyles.Number | NumberStyles.AllowCurrencySymbol | NumberStyles.AllowExponent;
        bool ok=double.TryParse(s,style,CultureInfo.InvariantCulture,out n) || double.TryParse(s,style,CultureInfo.CurrentCulture,out n);
        if(ok && percent) n/=100; return ok;
    }
    public static TypedResult Type(RawSheet sheet,int headerRow,IReadOnlyList<TypeFix> fixes,string? suppressColumn=null,string? suppressFix=null) {
        int header = headerRow - sheet.FirstRow; if(header<0 || header>=sheet.Rows) throw new ArgumentOutOfRangeException(nameof(headerRow));
        int count=sheet.Rows-header-1; var typed=new List<Column>(); var raw=new List<Column>(); var report=new CleanupReport();
        for(int c=0;c<sheet.Cols;c++) {
            string name=sheet.Text(header,c).Trim(); if(name.Length==0) name="Column " + (c+1);
            var original=new string?[count]; var cleaned=new string?[count]; int trimmed=0;
            for(int i=0;i<count;i++) { string v=sheet.Text(header+1+i,c); original[i]=v.Length==0?null:v; string t=string.Equals(name,suppressColumn,StringComparison.OrdinalIgnoreCase) && suppressFix is "trim" or "trimmed" ? v : v.Trim(); cleaned[i]=t.Length==0?null:t; if(v!=t) trimmed++; }
            raw.Add(TextColumn.From(name,original)); if(trimmed>0) report.Items.Add(new(name,"trimmed",trimmed,""));
            var rules=fixes.Where(f=>string.Equals(f.Column,name,StringComparison.OrdinalIgnoreCase)).ToList();
            foreach(var rule in rules) if(rule.Fix is "upper" or "padZeros" or "trim") { int changed=0; for(int i=0;i<count;i++) { var v=cleaned[i]; if(v==null) continue; string next=rule.Fix switch { "upper"=>v.ToUpperInvariant(), "padZeros" when v.All(char.IsDigit)=>v.PadLeft(rule.Width,'0'), _=>v.Trim() }; if(next!=v) changed++; cleaned[i]=next; } if(changed>0) report.Items.Add(new(name,rule.Fix,changed,"")); }
            var nonempty=Enumerable.Range(0,count).Where(i=>cleaned[i]!=null).ToArray();
            bool leading=nonempty.Any(i=>cleaned[i]!.Length>1 && cleaned[i]![0]=='0' && cleaned[i]!.All(char.IsDigit));
            if(string.Equals(name,suppressColumn,StringComparison.OrdinalIgnoreCase) && suppressFix=="leadingZerosKept") leading=false;
            if(leading) report.Items.Add(new(name,"leadingZerosKept",nonempty.Length,""));
            if(nonempty.Any(i=>cleaned[i]!.Any(char.IsUpper)) && nonempty.Any(i=>cleaned[i]!.Any(char.IsLower))) report.Items.Add(new(name,"mixedCase",nonempty.Length,""));
            var explicitType=rules.LastOrDefault(f=>f.Fix is "text" or "number" or "date");
            string? dateFormat=explicitType?.DateFormat;
            if(dateFormat==null) dateFormat=Formats.FirstOrDefault(fmt=>nonempty.Length>0 && nonempty.Count(i=>DateTime.TryParseExact(cleaned[i],fmt,CultureInfo.InvariantCulture,DateTimeStyles.None,out _)) >= Math.Ceiling(nonempty.Length*.95));
            int dateNumbers=nonempty.Count(i=>sheet.Kind(header+1+i,c)==CellKind.Number && sheet.IsDateFormat(header+1+i,c));
            int numeric=nonempty.Count(i=>sheet.Kind(header+1+i,c)==CellKind.Number || Numeric(cleaned[i]!,out _));
            bool isDate=explicitType?.Fix=="date" || (explicitType==null && !leading && nonempty.Length>0 && (dateNumbers>=Math.Ceiling(nonempty.Length*.95) || dateFormat!=null));
            bool isNumber=explicitType?.Fix=="number" || (explicitType==null && !leading && !isDate && nonempty.Length>0 && numeric>=Math.Ceiling(nonempty.Length*.95));
            if(isDate || isNumber) {
                var vals=new double[count]; var nulls=new bool[count]; int converted=0;
                for(int i=0;i<count;i++) { int row=header+1+i; if(cleaned[i]==null) {nulls[i]=true; continue;}
                    if(sheet.Kind(row,c)==CellKind.Number && (!isDate || sheet.IsDateFormat(row,c))) vals[i]=sheet.Num(row,c);
                    else if(isDate && dateFormat!=null && DateTime.TryParseExact(cleaned[i],dateFormat,CultureInfo.InvariantCulture,DateTimeStyles.None,out var dt)) {vals[i]=dt.ToOADate(); converted++;}
                    else if(isNumber && Numeric(cleaned[i]!,out var num)) {vals[i]=num; converted++;}
                    else nulls[i]=true;
                }
                typed.Add(new NumberColumn(name,vals,nulls.Any(x=>x)?nulls:null,isDate?ColType.Date:ColType.Number));
                if(converted>0) report.Items.Add(new(name,isDate?"dateFromText":"numberFromText",converted,""));
            } else typed.Add(TextColumn.From(name,cleaned));
        }
        var rowNumbers=Enumerable.Range(headerRow+1,count).Select(x=>(double)x).ToArray(); typed.Add(new NumberColumn("_r",rowNumbers,null)); raw.Add(new NumberColumn("_r",rowNumbers,null));
        return new TypedResult(new Table(sheet.Name,typed),report,new Table(sheet.Name,raw));
    }
}
