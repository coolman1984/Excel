using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using ExcelViewer.Core;
namespace ExcelApp.Data;
public static class ExportChild
{
    static void Release(object? item) {if(item!=null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item);}
    public static int Run(string jobPath) {
        StoreRoot? root=null;string? relative=null;dynamic? xl=null,book=null;int code=1;
        try {
            string full=Path.GetFullPath(jobPath);string testPrefix=@"D:\ExcelAppTest\";
            if(full.StartsWith(testPrefix,StringComparison.OrdinalIgnoreCase)) {string rel=Path.GetRelativePath(testPrefix,full);string folder=rel.Split(Path.DirectorySeparatorChar)[0];root=StoreRoot.ForTest(Path.Combine(testPrefix,folder));}
            else root=StoreRoot.Production();
            relative=Path.GetRelativePath(root.Folder,full).Replace('\\','/');if(!relative.StartsWith("exports/",StringComparison.OrdinalIgnoreCase) || !relative.EndsWith(".json",StringComparison.OrdinalIgnoreCase) || !string.Equals(root.PathOf(relative),full,StringComparison.OrdinalIgnoreCase))throw new UserError("Invalid export job.");
            var job=JsonSerializer.Deserialize<ExportJob>(root.ReadAllText(relative),Json.Options) ?? throw new UserError("Invalid export job.");
            if(job.Pdf) throw new UserError("PDF export is not approved.");
            var type=Type.GetTypeFromProgID("Excel.Application") ?? throw new UserError("Excel is unavailable.");xl=Activator.CreateInstance(type);xl.Visible=false;xl.DisplayAlerts=false;xl.EnableEvents=false;xl.AutomationSecurity=3;
            book=string.IsNullOrEmpty(job.Template)?xl.Workbooks.Add():xl.Workbooks.Open(job.Template,UpdateLinks:0,ReadOnly:true,Notify:false);
            foreach(var item in job.Sheets) {
                Console.WriteLine("PROGRESS Writing a sheet");dynamic? sheet=null,first=null,range=null;
                try {
                    try {sheet=book.Worksheets[item.Sheet];}catch(COMException) {sheet=book.Worksheets.Add();sheet.Name=item.Sheet;}
                    try {first=sheet.Range[item.StartCell];} catch(COMException) {first=book.Names.Item(item.StartCell).RefersToRange;}
                    int row=(int)first.Row,col=(int)first.Column;int n=item.Rows.Count+(item.Headers?1:0);int width=item.Columns.Count;
                    if(n>0 && width>0) {var data=new object[n,width];int offset=0;if(item.Headers){for(int c=0;c<width;c++)data[0,c]=item.Columns[c];offset=1;}
                        for(int r=0;r<item.Rows.Count;r++)for(int c=0;c<width;c++){string? value=c<item.Rows[r].Length?item.Rows[r][c]:null;data[r+offset,c]=ConvertValue(value);}
                        range=sheet.Range[sheet.Cells[row,col],sheet.Cells[row+n-1,col+width-1]];range.Value2=data;
                        if(item.Headers && string.IsNullOrEmpty(job.Template)) {dynamic h=sheet.Range[sheet.Cells[row,col],sheet.Cells[row,col+width-1]];try {h.Font.Bold=true;}finally {Release(h);}}
                        for(int c=0;c<Math.Min(width,item.NumberFormats.Count);c++)if(!string.IsNullOrEmpty(item.NumberFormats[c]) && item.Rows.Count>0) {dynamic column=sheet.Range[sheet.Cells[row+offset,col+c],sheet.Cells[row+n-1,col+c]];try {column.NumberFormat=item.NumberFormats[c];}finally {Release(column);}}
                        if(string.IsNullOrEmpty(job.Template)) range.Columns.AutoFit();
                    }
                } finally {Release(range);Release(first);Release(sheet);}
            }
            book.SaveAs(job.Target,FileFormat:51);code=0;
        } catch(Exception) {code=1;}
        finally {try {if(book!=null)book.Close(false);}catch(Exception) {code=1;}try {if(xl!=null)xl.Quit();}catch(Exception) {code=1;}Release(book);Release(xl);if(root!=null && relative!=null)try {root.Delete(relative);}catch(Exception) {}}
        Console.WriteLine(code==0?"RESULT ok":"RESULT error Excel could not export the workbook.");return code;
    }
    static object? ConvertValue(string? value) {if(value==null)return null;if(DateTime.TryParseExact(value,new[]{"yyyy-MM-dd","yyyy-MM-dd HH:mm:ss"},CultureInfo.InvariantCulture,DateTimeStyles.None,out var date))return date.ToOADate();if(!(value.Length>1 && value[0]=='0' && char.IsDigit(value[1])) && double.TryParse(value,NumberStyles.Float,CultureInfo.InvariantCulture,out double number))return number;return value;}
}
