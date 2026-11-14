// JSON message router between the interface and the session (plan 2.3).
using System.Diagnostics;
using System.Text.Json;

namespace ExcelApp;

/// <summary>A sentence meant for the user. Only our own text goes in here, never text from Excel.</summary>
public sealed class UserError(string message) : Exception(message) { }

public sealed class Bridge
{
    readonly Session session;
    readonly SemaphoreSlim serial = new(1, 1);
    readonly Dictionary<string, Func<JsonElement, Task<object?>>> handlers = new();
    // Commands that only read; they may run beside a long command.
    static readonly HashSet<string> Parallel = new() { "ready", "recent", "log", "cancelOpen", "cancelSave", "closeAnswer", "agent.status", "agent.stop" };
    public Action<string>? Post;
    static readonly JsonSerializerOptions Json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.AllowNamedFloatingPointLiterals };

    public Bridge(Session session)
    {
        this.session = session;
        session.Emit = (name, payload) => PostEvent(name, payload);
        session.Register(handlers);
    }

    /// <summary>Tells the interface something happened outside it (an AI agent changed the workbook).</summary>
    public void Notify(string name, object? payload) => PostEvent(name, payload);

    void PostEvent(string name, object? payload)
    {
        var d = new Dictionary<string, object?> { ["event"] = name };
        if (payload != null)
            foreach (var p in payload.GetType().GetProperties()) d[char.ToLowerInvariant(p.Name[0]) + p.Name[1..]] = p.GetValue(payload);
        Post?.Invoke(JsonSerializer.Serialize(d, Json));
    }

    public Task HandleMessage(string json) => Task.Run(async () => { var reply = await Dispatch(json); Post?.Invoke(reply); });

    /// <summary>Runs one request and returns the reply JSON. Used by the window and by tests.</summary>
    public async Task<string> Dispatch(string json)
    {
        int id = 0; string cmd = "?"; var sw = Stopwatch.StartNew(); bool locked = false;
        try
        {
            using var doc = JsonDocument.Parse(json); var root = doc.RootElement.Clone();
            id = root.TryGetProperty("id", out var i) ? i.GetInt32() : 0;
            cmd = root.GetProperty("cmd").GetString() ?? "";
            if (!handlers.TryGetValue(cmd, out var h)) { cmd = "unknown"; throw new UserError("That command is not known."); }
            if (!Parallel.Contains(cmd)) { await serial.WaitAsync(); locked = true; }
            if (session.Busy && !Parallel.Contains(cmd))
                throw new UserError("Please wait until the current work is finished.");
            object? result = await h(root);
            if (cmd != "rows" && cmd != "log") Log.Write($"cmd {cmd} {sw.ElapsedMilliseconds} ms");
            return JsonSerializer.Serialize(new { id, ok = true, result }, Json);
        }
        catch (Exception e)
        {
            Log.Write($"cmd {cmd} failed {e.GetType().Name}");
            string text = e is UserError ? e.Message : $"Something went wrong ({e.GetType().Name}). Your file is not changed.";
            return JsonSerializer.Serialize(new { id, ok = false, error = text }, Json);
        }
        finally { if (locked) serial.Release(); }
    }
}
