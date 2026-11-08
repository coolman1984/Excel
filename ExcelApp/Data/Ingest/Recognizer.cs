using System.Text.RegularExpressions;
namespace ExcelApp.Data;
public sealed record Recognition(int HeaderRow,double Score,IReadOnlyList<string> Missing,IReadOnlyList<string> Added);
public static class Recognizer
{
    public static string Normalize(string s) => Regex.Replace(s.Trim().ToLowerInvariant(),@"\s+"," ");
    public static bool Matches(string name,string pattern) {
        if(name.StartsWith("~$",StringComparison.OrdinalIgnoreCase) || new[]{".tmp",".crdownload",".partial",".part"}.Any(x=>name.EndsWith(x,StringComparison.OrdinalIgnoreCase))) return false;
        return Regex.IsMatch(name,"^"+Regex.Escape(pattern).Replace("\\*",".*").Replace("\\?",".")+"$",RegexOptions.IgnoreCase|RegexOptions.CultureInvariant);
    }
    public static Recognition Recognize(RawSheet sheet,IngestRecipe recipe) {
        var expected=recipe.HeaderSignature.Select(Normalize).ToHashSet(StringComparer.Ordinal);
        int start=recipe.HeaderRow>0?recipe.HeaderRow-sheet.FirstRow:0; int end=recipe.HeaderRow>0?start+1:Math.Min(sheet.Rows,20);
        if(start<0 || start>=sheet.Rows) throw new UserError("The header row is outside the sheet.");
        Recognition? best=null;
        for(int r=start;r<end;r++) {var actual=Enumerable.Range(0,sheet.Cols).Select(c=>Normalize(sheet.Text(r,c))).Where(x=>x.Length>0).ToHashSet(StringComparer.Ordinal);
            var missing=recipe.HeaderSignature.Where(x=>!actual.Contains(Normalize(x))).ToArray(); var added=Enumerable.Range(0,sheet.Cols).Select(c=>sheet.Text(r,c).Trim()).Where(x=>x.Length>0 && !expected.Contains(Normalize(x))).ToArray();
            double score=expected.Count==0?0:(double)(expected.Count-missing.Length)/expected.Count;
            if(best==null || score>best.Score) best=new Recognition(r+sheet.FirstRow,score,missing,added);
        } return best!;
    }
}
