using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelViewer.Core;

using ExcelApp;
public static class StyleTests
{
    public static void Run(Action<string,bool,string> check,bool offline)
    {
        var styles=new StyleSidecar();styles.Sheets[0]=new SheetStyles();styles.Sheets[0].Columns[2]=new(){new StyleRun(1,3,0xff0000,0x123456,true,2),new StyleRun(5,8,0x0000ff,null,null,1)};
        check("style lookup start",StyleStore.StyleAt(styles,0,1,2)?.Fill==0xff0000,"");
        check("style lookup end",StyleStore.StyleAt(styles,0,3,2)?.Bold==true,"");
        check("style lookup second run",StyleStore.StyleAt(styles,0,8,2)?.Fill==0x0000ff,"");
        check("style lookup gap",StyleStore.StyleAt(styles,0,4,2)==null,"");
        check("style lookup missing",StyleStore.StyleAt(styles,1,1,2)==null && StyleStore.StyleAt(styles,0,0,0)==null,"");
        check("style path extension",Path.GetExtension(CachePolicy.SidecarPathFor(@"D:\ExcelAppTest\syn-style.xlsx"))==".xvs","");
        OfflineRoundTrip(check);
        if(offline)return;
        if(Machine.Commit()>=Machine.StopAt){check("style live memory available",false,"memory");return;}
        Directory.CreateDirectory(@"D:\ExcelAppTest");string path=@"D:\ExcelAppTest\syn-styled.xlsx";
        dynamic? xl=null,wb=null;bool enabled=StyleStore.Enabled;int calls=StyleStore.MaxCalls;double seconds=StyleStore.MaxSeconds;
        try {
            xl=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!) ?? throw new InvalidOperationException();xl.Visible=false;xl.DisplayAlerts=false;xl.AutomationSecurity=3;xl.ScreenUpdating=false;xl.EnableEvents=false;
            wb=xl.Workbooks.Add();while(wb.Worksheets.Count>1)wb.Worksheets[wb.Worksheets.Count].Delete();dynamic ws=wb.Worksheets[1];
            Console.WriteLine("STYLE PHASE generating");
            // The measured gate: exactly 20 columns by 200,000 rows, values written in blocks.
            for(int row=1;row<=200000;row+=1000) {if(Machine.Commit()>=Machine.StopAt)throw new InvalidOperationException();var block=new object[1000,20];for(int r=0;r<1000;r++)for(int c=0;c<20;c++)block[r,c]=(double)(row+r+c);ws.Range[$"A{row}:T{row+999}"].Value2=block;}
            ws.Range["K1:O200000"].Interior.Color=65535;
            ws.Range["P1:R1"].Font.Bold=true;ws.Range["P1:R1"].Interior.Color=255;
            // Excel fills the two-row format pattern in one operation, preserving the numeric values.
            ws.Range["S2:T2"].Interior.Color=0xff0000;
            ws.Range["S1:T2"].AutoFill(ws.Range["S1:T200000"],3); // xlFillFormats
            wb.SaveAs(path,51);
            var fi=new FileInfo(path);var cache=new CacheData {SourceLength=fi.Length,SourceTicks=fi.LastWriteTimeUtc.Ticks};cache.Sheets.Add(new SheetInfo{FirstRow=1,FirstCol=1,Rows=200000,Cols=20});
            Console.WriteLine("STYLE PHASE measuring");
            var gateWatch=Stopwatch.StartNew();
            var side=StyleStore.Read((object)xl,(object)wb,cache);StyleStore.Write(path,cache,side);gateWatch.Stop();double measured=gateWatch.Elapsed.TotalSeconds;long measuredCalls=StyleStore.LastCalls;
            Console.WriteLine(FormattableString.Invariant($"STYLE GATE seconds={measured:F3} calls={measuredCalls} partial={side.Partial} errors={StyleStore.LastErrors} default={(measured>90?"off":"on")}"));
            check("style gate measured",StyleStore.LastErrors==0,"errors="+StyleStore.LastErrors);
            if(measured>=20 && measured<=90){StyleStore.MaxCalls=StyleStore.CallBudget/2;StyleStore.MaxSeconds=StyleStore.SecondsBudget/2;}
            check("style gate plain columns",!side.Sheets[0].Columns.ContainsKey(0),"");
            check("style gate fill runs",side.Sheets[0].Columns.TryGetValue(10,out var fill) && fill.Count==1 && fill[0].From==0 && fill[0].To==199999 && fill[0].Fill==0xffff00,"");
            check("style gate header runs",StyleStore.StyleAt(side,0,0,15)?.Bold==true && StyleStore.StyleAt(side,0,1,15)==null,"");
            var gateLoaded=StyleStore.Load(path,cache);
            check("style gate disk reload",gateLoaded!=null && Same(side,gateLoaded),"");
            check("style gate all plain columns",Enumerable.Range(0,10).All(c=>!side.Sheets[0].Columns.ContainsKey(c)),"");
            check("style gate all fill run lists",Enumerable.Range(10,5).All(c=>side.Sheets[0].Columns.TryGetValue(c,out var runs) && runs.SequenceEqual(new[]{new StyleRun(0,199999,0xffff00,null,null,null)})),"");
            check("style gate all header run lists",Enumerable.Range(15,3).All(c=>side.Sheets[0].Columns.TryGetValue(c,out var runs) && runs.SequenceEqual(new[]{new StyleRun(0,0,0xff0000,null,true,null)})),"");
            check("style gate known zebra runs",Enumerable.Range(18,2).All(c=>!side.Sheets[0].Columns.TryGetValue(c,out var runs) ? side.Partial : runs.All(r=>r.From==r.To && r.From%2==1 && r.Fill==0x0000ff && r.FontColor==null && r.Bold==null && r.Align==null)),"");
            // Small deterministic round trip, independent of the striped workbook's budget.
            ws.Cells.Clear();ws.Range["A1:D8"].Value2=7;ws.Range["B2:B4"].Interior.Color=255;ws.Range["B2:B4"].Font.Bold=true;ws.Range["B2:B4"].Font.Color=0x563412;ws.Range["B2:B4"].HorizontalAlignment=-4108;
            ws.Rows[3].Hidden=true;ws.Columns[3].Hidden=true;ws.Activate();wb.Windows[1].SplitRow=1;wb.Windows[1].SplitColumn=1;wb.Windows[1].FreezePanes=true;wb.Save();
            fi.Refresh();cache.SourceLength=fi.Length;cache.SourceTicks=fi.LastWriteTimeUtc.Ticks;cache.Sheets[0].Rows=8;cache.Sheets[0].Cols=4;
            StyleStore.Enabled=true;StyleStore.AfterRead((object)xl,(object)wb,cache);var loaded=StyleStore.Load(path,cache);
            check("style sidecar round trip",loaded!=null && !loaded.Partial,"");
            check("style borrowed worksheet valid",(int)ws.Index==1,"");
            if(loaded!=null) {
                check("style expected run list",loaded.Sheets[0].Columns[1].SequenceEqual(new[]{new StyleRun(1,3,0xff0000,0x123456,true,2)}),"");
                check("style hidden rows",loaded.Sheets[0].HiddenRows.SequenceEqual(new[]{(2,2)}),"runs="+loaded.Sheets[0].HiddenRows.Count);
                check("style hidden columns",loaded.Sheets[0].HiddenCols.SequenceEqual(new[]{(2,2)}),"runs="+loaded.Sheets[0].HiddenCols.Count);
                check("style frozen panes",loaded.Sheets[0].FrozenRows==1 && loaded.Sheets[0].FrozenCols==1,"");
            }
            cache.SourceTicks++;check("style stale rejected",StyleStore.Load(path,cache)==null,"");cache.SourceTicks--;
            StyleStore.MaxCalls=1;var partial=StyleStore.Read((object)xl,(object)wb,cache);check("style budget bounded",partial.Partial && StyleStore.LastCalls<=1,"");
            StyleStore.MaxCalls=StyleStore.CallBudget;ws.Cells.Clear();ws.Rows.Hidden=false;ws.Columns.Hidden=false;ws.Range["A1:D8"].Value2=7;wb.Windows[1].FreezePanes=false;
            wb.Save();
            double plainSeconds=0;
            var built=CacheBuilder.BuildFromOpenWorkbook((object)xl,(object)wb,path,CachePolicy.PathFor(path),_=>{},()=>false,(app,book,data)=>{
                var watch=Stopwatch.StartNew();StyleStore.AfterRead(app,book,data);watch.Stop();plainSeconds=watch.Elapsed.TotalSeconds;
            });
            var actualCache=CacheData.Load(CachePolicy.PathFor(path));var plain=StyleStore.Load(path,actualCache);
            check("style plain cache and sidecar",built.Code==0 && plain!=null,"");
            check("style plain near empty",plain!=null && plain.Sheets[0].Columns.Count==0 && !plain.Partial,"");
            Console.WriteLine(FormattableString.Invariant($"STYLE PLAIN seconds={plainSeconds:F3} bytes={new FileInfo(CachePolicy.SidecarPathFor(path)).Length}"));
            check("style plain under one second",plainSeconds<1,"");
        } catch(Exception e){check("style live exception",false,e.GetType().Name);}
        finally {
            StyleStore.Enabled=enabled;StyleStore.MaxCalls=calls;StyleStore.MaxSeconds=seconds;
            try{if(wb!=null)wb.Close(false);}finally{if(xl!=null)xl.Quit();}if(wb!=null)Marshal.FinalReleaseComObject(wb);if(xl!=null)Marshal.FinalReleaseComObject(xl);
            CachePolicy.Delete(CachePolicy.PathFor(path));CachePolicy.Delete(CachePolicy.PathFor(path)+".part");
            CachePolicy.Delete(CachePolicy.SidecarPathFor(path));CachePolicy.Delete(CachePolicy.SidecarPathFor(path)+".part");if(File.Exists(path))File.Delete(path);
        }
    }
    static bool Same(StyleSidecar a,StyleSidecar b) => a.Partial==b.Partial && a.Sheets.Count==b.Sheets.Count && a.Sheets.All(pair=>
        b.Sheets.TryGetValue(pair.Key,out var other) && pair.Value.FrozenRows==other.FrozenRows && pair.Value.FrozenCols==other.FrozenCols &&
        pair.Value.HiddenRows.SequenceEqual(other.HiddenRows) && pair.Value.HiddenCols.SequenceEqual(other.HiddenCols) &&
        pair.Value.Columns.Count==other.Columns.Count && pair.Value.Columns.All(col=>other.Columns.TryGetValue(col.Key,out var runs) && col.Value.SequenceEqual(runs)));
    static void OfflineRoundTrip(Action<string,bool,string> check)
    {
        bool refused=false;
        try {using var denied=CachePolicy.OpenForWrite(@"D:\ExcelAppTest\syn-outside.xvs");}
        catch(CachePolicyException){refused=true;}
        check("style outside write refused",refused,"");
        Directory.CreateDirectory(@"D:\ExcelAppTest");
        string source=@"D:\ExcelAppTest\syn-style-offline-"+Guid.NewGuid().ToString("N")+".fixture";
        // A metadata-only source fixture, not a workbook. No workbook is created or parsed outside Excel.
        using(var created=new FileStream(source,FileMode.CreateNew,FileAccess.Write)){}
        var info=new FileInfo(source);var cache=new CacheData{SourceLength=info.Length,SourceTicks=info.LastWriteTimeUtc.Ticks};
        cache.Sheets.Add(new SheetInfo{Rows=12,Cols=4});
        var expected=new StyleSidecar{Partial=true};expected.Sheets[0]=new SheetStyles{FrozenRows=1,FrozenCols=2,HiddenRows=new(){(2,3)},HiddenCols=new(){(0,0)}};
        expected.Sheets[0].Columns[2]=new(){new StyleRun(1,3,0x102030,0xabcdef,true,2),new StyleRun(5,8,null,null,false,0)};
        try {
            check("style missing sidecar",StyleStore.Load(source,cache)==null,"");
            StyleStore.Write(source,cache,expected);var actual=StyleStore.Load(source,cache);
            check("style handmade disk round trip",actual!=null && Same(expected,actual),"");
            check("style explicit false retained",actual!=null && StyleStore.StyleAt(actual,0,6,2)?.Bold==false,"");
            check("style atomic no part",!File.Exists(CachePolicy.SidecarPathFor(source)+".part"),"");
            cache.SourceTicks++;check("style stale cache rejected",StyleStore.Load(source,cache)==null,"");cache.SourceTicks--;
            File.SetLastWriteTimeUtc(source,info.LastWriteTimeUtc.AddSeconds(2));info.Refresh();cache.SourceTicks=info.LastWriteTimeUtc.Ticks;
            check("style stale sidecar header rejected",StyleStore.Load(source,cache)==null,"");
            using(var corrupt=CachePolicy.OpenForWrite(CachePolicy.SidecarPathFor(source)))corrupt.Write(new byte[]{0,1,2});
            check("style damaged sidecar rejected",StyleStore.Load(source,cache)==null,"");
        }
        finally {CachePolicy.Delete(CachePolicy.SidecarPathFor(source));CachePolicy.Delete(CachePolicy.SidecarPathFor(source)+".part");File.Delete(source);}
    }

}
