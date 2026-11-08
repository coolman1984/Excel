using System.Diagnostics;
using System.Text.Json;
namespace ExcelApp.Data;
public sealed class Exporter : IExporter
{
    readonly StoreRoot root; readonly WorkQueue queue; readonly Func<Settings> settings;
    public Exporter(StoreRoot root,WorkQueue queue,Func<Settings> settings) {this.root=root;this.queue=queue;this.settings=settings;foreach(var file in root.List("exports","*.json"))root.Delete(file);}
    public Task ExportAsync(ExportJob job,IProgress<string>? progress,CancellationToken token) {
        if(job.Pdf && !settings().AllowPdf) throw new UserError("PDF export is not approved.");
        var completion=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        queue.Enqueue("export",async queueToken=>{try {using var linked=CancellationTokenSource.CreateLinkedTokenSource(token,queueToken);await Run(job,progress,linked.Token);completion.TrySetResult(true);}catch(OperationCanceledException){completion.TrySetCanceled();}catch(Exception e){completion.TrySetException(e);}});return completion.Task;
    }
    async Task Run(ExportJob job,IProgress<string>? progress,CancellationToken token) {
        string relative=$"exports/{Guid.NewGuid():N}.json";root.WriteAllText(relative,JsonSerializer.Serialize(job,Json.Options));
        var psi=new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory,"ExcelApp.exe")) {UseShellExecute=false,RedirectStandardOutput=true,RedirectStandardInput=true,CreateNoWindow=true};psi.ArgumentList.Add("--export");psi.ArgumentList.Add(root.PathOf(relative));
        using var child=Process.Start(psi) ?? throw new UserError("Excel could not be started.");
        var read=Task.Run(async()=>{string? line;string? result=null;while((line=await child.StandardOutput.ReadLineAsync())!=null) {if(line.StartsWith("PROGRESS ",StringComparison.Ordinal)) progress?.Report(line[9..]);else if(line.StartsWith("RESULT ",StringComparison.Ordinal))result=line;}return result;});
        var exited=child.WaitForExitAsync(CancellationToken.None);var finished=await Task.WhenAny(exited,Task.Delay(TimeSpan.FromMinutes(10),token));
        if(finished!=exited) {token.ThrowIfCancellationRequested();throw new UserError("Excel is still busy.");}
        await exited;string? answer=await read;if(child.ExitCode!=0 || answer!="RESULT ok") throw new UserError("Excel could not export the workbook.");
    }
}
