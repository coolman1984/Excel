// The running app's door for AI agents: a named pipe that only this Windows account can open. Each line is one interface
// command ({"id":1,"cmd":"agent.read",...}); the answer is the same JSON the interface gets. ExcelApp.exe --mcp connects here.
using System.IO.Pipes;
using System.Text;
using System.Text.Json;

namespace ExcelApp.Agent;

public sealed class AgentLink : IDisposable
{
    /// <summary>One name per Windows account. PipeOptions.CurrentUserOnly keeps other accounts out.</summary>
    public static string DefaultPipe => "ExcelApp.Agent." + new string(Environment.UserName.Select(c => char.IsLetterOrDigit(c) ? c : '_').ToArray());
    // Commands that only look: the interface does not need to redraw after them.
    static readonly HashSet<string> Reads = new(StringComparer.Ordinal) { "ready", "recent", "log", "workbook", "rows", "cellDetail", "stats", "find", "filterValues", "info", "edge", "locate", "agent.read", "recovery" };
    const int MaxLine = 16 << 20;

    readonly Func<string, Task<string>> dispatch; readonly Action<string>? changed; readonly string name;
    readonly CancellationTokenSource stop = new();
    public AgentLink(Func<string, Task<string>> dispatch, Action<string>? changed = null, string? pipeName = null)
    { this.dispatch = dispatch; this.changed = changed; name = pipeName ?? DefaultPipe; }

    public void Start() => _ = Task.Run(Loop);

    async Task Loop()
    {
        while (!stop.IsCancellationRequested)
        {
            NamedPipeServerStream server;
            try { server = new NamedPipeServerStream(name, PipeDirection.InOut, NamedPipeServerStream.MaxAllowedServerInstances, PipeTransmissionMode.Byte, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { Log.Write("agent link " + e.GetType().Name); try { await Task.Delay(5000, stop.Token); } catch (OperationCanceledException) { return; } continue; }
            try { await server.WaitForConnectionAsync(stop.Token); }
            catch (OperationCanceledException) { server.Dispose(); return; }
            catch (IOException) { server.Dispose(); continue; }
            _ = Task.Run(() => Serve(server));
        }
    }

    async Task Serve(NamedPipeServerStream pipe)
    {
        try
        {
            using (pipe)
            using (var reader = new StreamReader(pipe, Encoding.UTF8, false, 1 << 16, true))
            using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 1 << 16, true) { AutoFlush = true })
            {
                string? line;
                while ((line = await reader.ReadLineAsync(stop.Token)) != null)
                {
                    if (line.Length == 0) continue;
                    string reply = line.Length > MaxLine ? "{\"id\":0,\"ok\":false,\"error\":\"The request is too large.\"}" : await dispatch(line);
                    await writer.WriteLineAsync(reply);
                    if (changed != null && Changed(line, reply) is string cmd) try { changed(cmd); } catch (Exception) { }
                }
            }
        }
        catch (Exception e) when (e is IOException or OperationCanceledException or ObjectDisposedException) { }
    }

    /// <summary>The command name when it succeeded and may have changed something the interface shows.</summary>
    internal static string? Changed(string request, string reply)
    {
        try
        {
            using var q = JsonDocument.Parse(request); using var r = JsonDocument.Parse(reply);
            string? cmd = q.RootElement.TryGetProperty("cmd", out var c) ? c.GetString() : null;
            bool ok = r.RootElement.TryGetProperty("ok", out var o) && o.ValueKind == JsonValueKind.True;
            return ok && cmd != null && !Reads.Contains(cmd) ? cmd : null;
        }
        catch (JsonException) { return null; }
    }

    public void Dispose() { stop.Cancel(); }
}

/// <summary>The MCP server's side of the door: one connection per call.</summary>
public interface IAppLink { JsonElement Call(string cmd, JsonElement args); }

public sealed class AppClient(string? pipeName = null, int connectMs = 800) : IAppLink
{
    readonly string name = pipeName ?? AgentLink.DefaultPipe;
    public JsonElement Call(string cmd, JsonElement args)
    {
        using var pipe = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.CurrentUserOnly);
        try { pipe.Connect(connectMs); }
        catch (Exception e) when (e is TimeoutException or IOException) { throw new UserError("ExcelApp is not open. Open ExcelApp (and a workbook), then try again."); }
        var request = new Dictionary<string, object?> { ["id"] = 1, ["cmd"] = cmd };
        if (args.ValueKind == JsonValueKind.Object) foreach (var p in args.EnumerateObject()) if (p.Name is not ("id" or "cmd")) request[p.Name] = p.Value;
        using var writer = new StreamWriter(pipe, new UTF8Encoding(false), 1 << 16, true) { AutoFlush = true };
        using var reader = new StreamReader(pipe, Encoding.UTF8, false, 1 << 16, true);
        writer.WriteLine(JsonSerializer.Serialize(request));
        string reply = reader.ReadLine() ?? throw new UserError("ExcelApp closed the connection.");
        using var doc = JsonDocument.Parse(reply);
        var root = doc.RootElement;
        if (!(root.TryGetProperty("ok", out var ok) && ok.ValueKind == JsonValueKind.True))
            throw new UserError(root.TryGetProperty("error", out var e) ? e.GetString() ?? "ExcelApp refused the command." : "ExcelApp refused the command.");
        return root.TryGetProperty("result", out var result) ? result.Clone() : default;
    }
}
