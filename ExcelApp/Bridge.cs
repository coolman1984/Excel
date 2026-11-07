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
    static readonly HashSet<string> Parallel = new() { "ready", "recent", "log", "cancelOpen", "cancelSave", "closeAnswer" };
    public Action<string>? Post;
    static readonly JsonSerializerOptions Json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.AllowNamedFloatingPointLiterals };

    public Bridge(Session session)
    {
        this.session = session;
        session.Emit = (name, payload) => PostEvent(name, payload);
        session.Register(handlers);
    }

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
            if (session.Busy && cmd is not ("ready" or "recent" or "log" or "cancelOpen" or "cancelSave" or "closeAnswer"))
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

/// <summary>Reading arguments out of a request.</summary>
public static class Arg
{
    public static int Int(JsonElement e, string name, int? fallback = null)
    {
        if (e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number) return (int)v.GetDouble();
        return fallback ?? throw new UserError("The request is missing a number.");
    }
    public static double Num(JsonElement e, string name, double? fallback = null)
    {
        if (e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number) return v.GetDouble();
        return fallback ?? throw new UserError("The request is missing a number.");
    }
    public static string? Str(JsonElement e, string name) => e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;
    public static bool Bool(JsonElement e, string name, bool fallback = false) => e.TryGetProperty(name, out var v) && v.ValueKind is JsonValueKind.True or JsonValueKind.False ? v.GetBoolean() : fallback;
    public static JsonElement? Obj(JsonElement e, string name) => e.TryGetProperty(name, out var v) && v.ValueKind is not (JsonValueKind.Null or JsonValueKind.Undefined) ? v : null;
    public static List<(int R1, int C1, int R2, int C2)> Ranges(JsonElement e, string name = "ranges")
    {
        var list = new List<(int, int, int, int)>();
        if (!e.TryGetProperty(name, out var arr) || arr.ValueKind != JsonValueKind.Array) return list;
        foreach (var r in arr.EnumerateArray()) list.Add(Rect(r));
        return list;
    }
    public static (int R1, int C1, int R2, int C2) Rect(JsonElement r)
    {
        int a = r[0].GetInt32(), b = r[1].GetInt32(), c = r[2].GetInt32(), d = r[3].GetInt32();
        return (Math.Min(a, c), Math.Min(b, d), Math.Max(a, c), Math.Max(b, d));
    }
}
