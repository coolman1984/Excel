// Synthetic COM round trips. Live fixtures belong only in D:\ExcelAppTest and start with syn.
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using ExcelViewer.Core;

using ExcelApp;

public static class SaveTests
{
    public static void Run(Action<string, bool, string> check, bool offline)
    {
        var output = new StringWriter(CultureInfo.InvariantCulture);
        int stopped = SaveBack.Run(new StringReader("{}"), output, 0);
        check("save memory guard", stopped == 90 && output.ToString().Contains("STOPPED memory") && output.ToString().Contains("RESULT code=90") && !output.ToString().Contains("OPENING"), "");
        output = new StringWriter();
        check("save malformed JSON", SaveBack.Run(new StringReader("{"), output) == 1 && !output.ToString().Contains("OPENING"), "");
        output = new StringWriter();
        string invalid = JsonSerializer.Serialize(new { source = @"D:\ExcelAppTest\syn-source.xlsx", target = @"D:\ExcelAppTest\syn-other.xlsx", sameFile = true });
        check("save rejects inconsistent target", SaveBack.Run(new StringReader(invalid), output) == 1 && !output.ToString().Contains("OPENING"), "");
        output = new StringWriter();
        invalid = JsonSerializer.Serialize(new { source = @"D:\ExcelAppTest\syn-source.csv", target = @"D:\ExcelAppTest\syn-source.csv", sameFile = true });
        check("review: child refuses same-file CSV saves", SaveBack.Run(new StringReader(invalid), output) == 1 && !output.ToString().Contains("OPENING"), "");
        output = new StringWriter();
        invalid = JsonSerializer.Serialize(new { source = @"D:\ExcelAppTest\syn-missing.xlsx", target = @"D:\ExcelAppTest\syn-missing.xlsx", sameFile = true, expectedLength = 1, expectedTicks = 1 });
        check("review: stale source refused before Excel starts", SaveBack.Run(new StringReader(invalid), output) == 7 && !output.ToString().Contains("OPENING"), "");
        if (offline) return;
        if (Machine.Commit() >= Machine.StopAt) { check("live save memory available", false, "memory"); return; }
        string root = @"D:\ExcelAppTest"; Directory.CreateDirectory(root);
        string id = Guid.NewGuid().ToString("N")[..8];
        string source = Path.Combine(root, "syn-save-" + id + ".xlsx"), target = Path.Combine(root, "syn-saved-" + id + ".xlsx"), structTarget = Path.Combine(root, "syn-struct-" + id + ".xlsx"), protectedPath = Path.Combine(root, "syn-protected-" + id + ".xlsx");
        try
        {
            Make(source, false); Make(protectedPath, true);
            object? before1 = null, before2 = null;
            WithBook(source, (xl, wb) => { before1 = wb.Worksheets[2].UsedRange.Value2; before2 = wb.Worksheets[3].UsedRange.Value2; });
            var built = CacheBuilder.Build(source, CachePolicy.PathFor(source), _ => { }, () => false);
            check("save initial cache build", built.Code == 0, "code=" + built.Code);
            if (built.Code == 0) {
                var initial = CacheData.Load(CachePolicy.PathFor(source));
                check("core row list subset", SortEngine.Sort(initial, initial.Sheets[0], 1, false, new[]{1,2}, default).SequenceEqual(new[]{2,1}), "");
                check("core row list stable ties", SortEngine.Sort(initial, initial.Sheets[1], 2, true, new[]{2,1}, default).SequenceEqual(new[]{2,1}), "");
            }
            // Construct the overlay first, then serialize absolute coordinates just as the host does.
            var set = new EditSet();
            set.Apply(0, new (int Row, int Col, CellEdit? After)[] {
                (1, 1, new CellEdit(EditKind.Number, Num:25)),
                (1, 0, new CellEdit(EditKind.Text, Text:"=synthetic literal")),
                (1, 3, new CellEdit(EditKind.Formula, Formula:"=SUM(B2:B3)")),
                (2, 0, new CellEdit(EditKind.Clear)),
                (0, 0, new CellEdit(EditKind.Text, Text:"synthetic header", Style:new StyleDelta(Bold:true, Fill:0xff0000, FontColor:0x123456, Italic:true, Underline:true, Align:2, NumberFormat:"@", Border:"bottom"))),
                (2, 4, new CellEdit(EditKind.Bool, Bool:true))
            });
            SaveBack.Job Job(string src, string dst, bool same) => new() {
                Source=src, Target=dst, SameFile=same,
                Edits=set.All.Select(pair => new SaveBack.Edit { Sheet="syn-one", Row=pair.Key.Row+1, Col=pair.Key.Col+1, Kind=pair.Value.Kind, Num=pair.Value.Num, Text=pair.Value.Text, Formula=pair.Value.Formula, Bool=pair.Value.Bool, Style=pair.Value.Style }).ToList(),
                ColWidths=new(){new SaveBack.Width { Sheet="syn-one", Col=2, Px=145 }},
                Freeze=new(){new SaveBack.Freeze { Sheet="syn-one", Rows=1, Cols=1 }}
            };
            foreach (bool same in new[] { false, true }) {
                var job = Job(same ? target : source, target, same); output = new StringWriter();
                int code = SaveBack.Run(new StringReader(JsonSerializer.Serialize(job)), output);
                string label = same ? "same" : "as";
                check("save " + label + " result", code == 0, "code=" + code);
                check("save " + label + " progress", InOrder(output.ToString(), "OPENING", "OPENED", "WRITING", "SAVING", "SAVED", "FORMATS", "READ", "CHECKING", "RESULT"), "");
                if (code != 0) continue;
                WithBook(target, (xl, wb) => {
                    dynamic ws = wb.Worksheets[1];
                    check("save " + label + " number", (double)ws.Cells[2,2].Value2 == 25, "");
                    check("save " + label + " recalculation", (double)ws.Cells[2,3].Value2 == 50, "");
                    check("save " + label + " formula", (string)ws.Cells[2,4].Formula == "=SUM(B2:B3)" && (double)ws.Cells[2,4].Value2 == 30, "");
                    check("save " + label + " clear", ws.Cells[3,1].Value2 == null, "");
                    check("save " + label + " literal text", (string)ws.Cells[2,1].Value2 == "=synthetic literal" && !(bool)ws.Cells[2,1].HasFormula, "");
                    check("save " + label + " boolean", (bool)ws.Cells[3,5].Value2, "");
                    check("save " + label + " fill and bold", (bool)ws.Cells[1,1].Font.Bold && Convert.ToInt32(ws.Cells[1,1].Interior.Color) == 255, "");
                    check("save " + label + " font and alignment", Convert.ToInt32(ws.Cells[1,1].Font.Color) == 0x563412 && (bool)ws.Cells[1,1].Font.Italic && Convert.ToInt32(ws.Cells[1,1].Font.Underline) == 2 && Convert.ToInt32(ws.Cells[1,1].HorizontalAlignment) == -4108, "");
                    check("save " + label + " format and border", (string)ws.Cells[1,1].NumberFormat == "@" && Convert.ToInt32(ws.Cells[1,1].Borders[9].LineStyle) == 1, "");
                    check("save " + label + " width", Math.Abs((double)ws.Columns[2].ColumnWidth - 20) < .2, "");
                    ws.Activate(); check("save " + label + " freeze", (bool)wb.Windows[1].FreezePanes && (int)wb.Windows[1].SplitRow == 1 && (int)wb.Windows[1].SplitColumn == 1, "");
                    check("save " + label + " untouched sheets", Equal(before1, (object)wb.Worksheets[2].UsedRange.Value2) && Equal(before2, (object)wb.Worksheets[3].UsedRange.Value2), "");
                });
                using (var stream = File.OpenRead(target)) { var bytes = new byte[8]; stream.ReadExactly(bytes); check("save " + label + " NASCA mark", Encoding.ASCII.GetString(bytes) == "<## NASC", ""); }
                var cache = CacheData.Load(CachePolicy.PathFor(target));
                check("save " + label + " cache fresh", cache.IsFreshFor(target) && CacheData.HeaderMatches(cache.Path,target), "");
                using (var reader = new SheetReader(cache,cache.Sheets[0])) check("save " + label + " cache values", reader.Cell(1,1).Text == "25" && reader.Cell(1,2).Text == "50", "");
                check("save " + label + " no part", !File.Exists(CachePolicy.PathFor(target)+".part") && !File.Exists(CachePolicy.SidecarPathFor(target)+".part"), "");
            }
            // Inserted and deleted rows and columns are replayed in Excel first; the cells are then written at the new addresses.
            var sjob = new SaveBack.Job { Source = source, Target = structTarget, SameFile = false,
                Structure = new() { new SaveBack.Structure { Sheet = "syn-one", Rows = true, At = 2, Count = 1, Insert = true }, new SaveBack.Structure { Sheet = "syn-one", Rows = false, At = 5, Count = 1, Insert = false } },
                Edits = new() { new SaveBack.Edit { Sheet = "syn-one", Row = 2, Col = 1, Kind = EditKind.Text, Text = "inserted" } } };
            output = new StringWriter();
            int scode = SaveBack.Run(new StringReader(JsonSerializer.Serialize(sjob)), output);
            check("save structure result", scode == 0, "code=" + scode);
            if (scode == 0) WithBook(structTarget, (xl, wb) => {
                dynamic ws = wb.Worksheets[1];
                check("save structure: the inserted row holds the new cell", (string)ws.Cells[2, 1].Value2 == "inserted", "");
                check("save structure: the old row moved down and Excel moved its formula", (string)ws.Cells[3, 1].Value2 == "text" && (string)ws.Cells[3, 3].Formula == "=B3*2", (string)ws.Cells[3, 3].Formula);
                check("save structure: the deleted column is gone", ws.Cells[1, 5].Value2 == null && (string)ws.Cells[1, 4].Value2 == "money", "");
            });
            byte[] hash = System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(source));
            var attr = File.GetAttributes(source);
            try {
                File.SetAttributes(source,attr | FileAttributes.ReadOnly); output=new StringWriter();
                check("save readonly", SaveBack.Run(new StringReader(JsonSerializer.Serialize(Job(source,source,true))),output)==4 && output.ToString().Contains("STOPPED readonly"), "");
                check("save readonly unchanged", hash.SequenceEqual(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(source))), "");
                output=new StringWriter();
                check("save readonly allows Save As", SaveBack.Run(new StringReader(JsonSerializer.Serialize(Job(source,target,false))),output)==0, "");
            } finally { File.SetAttributes(source,attr); }
            output=new StringWriter();
            check("save protected", SaveBack.Run(new StringReader(JsonSerializer.Serialize(Job(protectedPath,protectedPath,true))),output)==6 && output.ToString().Contains("STOPPED protected"), "");
        }
        catch (Exception e) { check("save live exception",false,e.GetType().Name); }
        finally {
            foreach (string path in new[]{source,target,structTarget,protectedPath}) { CachePolicy.Delete(CachePolicy.PathFor(path)); CachePolicy.Delete(CachePolicy.PathFor(path)+".part"); CachePolicy.Delete(CachePolicy.SidecarPathFor(path)); CachePolicy.Delete(CachePolicy.SidecarPathFor(path)+".part"); if(File.Exists(path))File.Delete(path); }
        }
    }
    static bool InOrder(string text, params string[] tokens) { int at=0; foreach(var token in tokens) { int next=text.IndexOf(token,at,StringComparison.Ordinal); if(next<0)return false; at=next+token.Length; }return true; }
    static bool Equal(object? a, object? b) {
        if(a is Array aa && b is Array bb) return aa.Rank==bb.Rank && Enumerable.Range(0,aa.Rank).All(i=>aa.GetLength(i)==bb.GetLength(i)) && aa.Cast<object?>().SequenceEqual(bb.Cast<object?>());
        return Equals(a,b);
    }
    internal static void WithBook(string path, Action<dynamic,dynamic> action) {
        dynamic? xl=null, wb=null;
        try { xl=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!) ?? throw new InvalidOperationException(); xl.Visible=false;xl.DisplayAlerts=false;xl.AutomationSecurity=3;xl.EnableEvents=false; wb=xl.Workbooks.Open(path,UpdateLinks:0,ReadOnly:true,Password:""); action(xl,wb); }
        finally { try {if(wb!=null)wb.Close(false);} finally {if(xl!=null)xl.Quit();} if(wb!=null)Marshal.FinalReleaseComObject(wb);if(xl!=null)Marshal.FinalReleaseComObject(xl); }
    }
    static void Make(string path,bool protect) {
        dynamic? xl=null,wb=null;
        try {
            xl=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!) ?? throw new InvalidOperationException();xl.Visible=false;xl.DisplayAlerts=false;xl.AutomationSecurity=3;xl.EnableEvents=false;
            wb=xl.Workbooks.Add();while(wb.Worksheets.Count<3)wb.Worksheets.Add();while(wb.Worksheets.Count>3)wb.Worksheets[wb.Worksheets.Count].Delete();
            for(int i=1;i<=3;i++) { dynamic ws=wb.Worksheets[i];ws.Name=i==1?"syn-one":"syn-"+i;ws.Range["A1:E3"].Value2=new object[,] {{"synthetic", "number", "formula", "money", "bool"},{"text",10,0,12.5,false},{"clear",5,0,24.5,true}};ws.Range["D2:D3"].NumberFormat="$0.00"; }
            dynamic first=wb.Worksheets[1]; first.Cells[2,3].Formula="=B2*2";
            if(protect)first.Protect();wb.SaveAs(path,51);
        } finally {try{if(wb!=null)wb.Close(false);}finally{if(xl!=null)xl.Quit();}if(wb!=null)Marshal.FinalReleaseComObject(wb);if(xl!=null)Marshal.FinalReleaseComObject(xl);}
    }
}
