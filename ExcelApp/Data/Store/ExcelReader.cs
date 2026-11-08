using System.Diagnostics;
using ExcelViewer.Core;
namespace ExcelApp.Data;

public sealed class RawSheet
{
    readonly byte[,] kinds; readonly double[,] nums; readonly string[,] texts; readonly bool[,] dates;
    public int Rows => kinds.GetLength(0); public int Cols => kinds.GetLength(1);
    public int FirstRow { get; init; } = 1; public int FirstCol { get; init; } = 1;
    public string Name { get; init; } = "";
    public RawSheet(byte[,] kinds, double[,] nums, string[,] texts, bool[,] dates) {
        this.kinds=kinds;
        if (nums.GetLength(0) != Rows || nums.GetLength(1) != Cols || texts.GetLength(0) != Rows || texts.GetLength(1) != Cols || dates.GetLength(0) != Rows || dates.GetLength(1) != Cols) throw new ArgumentException("Sheet dimensions differ.");
        this.nums=nums; this.texts=texts; this.dates=dates;
    }
    public byte Kind(int r,int c) => kinds[r,c]; public double Num(int r,int c) => nums[r,c];
    public string Text(int r,int c) => texts[r,c] ?? ""; public bool IsDateFormat(int r,int c) => dates[r,c];
}
public interface IRawSource { RawSheet ReadSheet(string file, string? sheet, CancellationToken token); IReadOnlyList<RawSheet> ReadSheets(string file,CancellationToken token) => new[] {ReadSheet(file,null,token)}; }
public sealed class ExcelReader : IRawSource
{
    public RawSheet ReadSheet(string file, string? sheet, CancellationToken token) => ReadInternal(file,sheet,false,token).Single();
    public IReadOnlyList<RawSheet> ReadSheets(string file,CancellationToken token) => ReadInternal(file,null,true,token);
    IReadOnlyList<RawSheet> ReadInternal(string file,string? sheet,bool all,CancellationToken token) {
        var psi = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "ExcelApp.exe")) { UseShellExecute=false, RedirectStandardOutput=true, RedirectStandardInput=true, CreateNoWindow=true };
        psi.ArgumentList.Add("--build"); psi.ArgumentList.Add(file);
        using var process = Process.Start(psi) ?? throw new UserError("Excel could not be started.");
        var output = process.StandardOutput.ReadToEndAsync();
        if (!process.WaitForExit((int)TimeSpan.FromMinutes(5).TotalMilliseconds)) throw new UserError("Excel is still busy.");
        token.ThrowIfCancellationRequested();
        string outputText = output.GetAwaiter().GetResult();
        if (process.ExitCode != 0 || !outputText.Split('\n').Any(line => line.Trim().StartsWith("RESULT code=0", StringComparison.Ordinal))) throw new UserError("Excel could not read the file.");
        var cache = CacheData.Load(CachePolicy.PathFor(file));
        var chosen = all ? cache.Sheets.Where(x=>!x.Hidden).ToArray() : new[] { string.IsNullOrEmpty(sheet) ? cache.Sheets.FirstOrDefault(x => !x.Hidden) : cache.Sheets.FirstOrDefault(x => string.Equals(x.Name,sheet,StringComparison.OrdinalIgnoreCase)) };
        if (chosen.Any(x=>x==null)) throw new UserError("The requested sheet was not found.");
        var result=new List<RawSheet>();
        foreach(var info in chosen) {
            var kinds = new byte[info!.Rows,info.Cols]; var nums = new double[info.Rows,info.Cols]; var texts = new string[info.Rows,info.Cols]; var dates = new bool[info.Rows,info.Cols];
            using var reader = new SheetReader(cache,info);
            for (int r=0;r<info.Rows;r++) { token.ThrowIfCancellationRequested(); for (int c=0;c<info.Cols;c++) {
                var chunk=reader.Column(r,c); int i=r%info.BlockRows; kinds[r,c]=chunk.KindAt(i);
                if (kinds[r,c] is CellKind.Number or CellKind.Bool or CellKind.Error) nums[r,c]=chunk.Nums[i];
                texts[r,c]=reader.Show(chunk,i,r,c).Text;
                int fmt=info.FormatAt(r,c); dates[r,c]=fmt>=0 && cache.Formats[fmt].Compiled(cache).IsDateFormat;
            } }
            result.Add(new RawSheet(kinds,nums,texts,dates) { Name=info.Name, FirstRow=info.FirstRow, FirstCol=info.FirstCol });
        }
        return result;
    }
}
