using System.Threading.Channels;
namespace ExcelApp.Data;
public sealed class WorkQueue : IDisposable
{
    readonly Channel<(string Label, Func<CancellationToken,Task> Work)> queue = Channel.CreateUnbounded<(string,Func<CancellationToken,Task>)>(new UnboundedChannelOptions { SingleReader=true });
    readonly CancellationTokenSource stop = new(); readonly Task worker; readonly object gate = new(); int pending; string? current;
    public WorkQueue() { worker=Task.Run(Run); }
    public void Enqueue(string label, Func<CancellationToken,Task> work) { lock(gate) { if(!queue.Writer.TryWrite((label,work))) throw new InvalidOperationException("Queue is closed."); pending++; } }
    public (int Pending,string? Current) Status { get { lock(gate) return (pending,current); } }
    async Task Run() { try { await foreach(var item in queue.Reader.ReadAllAsync(stop.Token)) { lock(gate) { pending--; current=item.Label; } try { await item.Work(stop.Token); } catch (OperationCanceledException) when(stop.IsCancellationRequested) { } catch (Exception) { } finally { lock(gate) current=null; } } } catch(OperationCanceledException) { } }
    public void Dispose() { queue.Writer.TryComplete(); stop.Cancel(); try { worker.GetAwaiter().GetResult(); } catch(OperationCanceledException) { } stop.Dispose(); }
}
