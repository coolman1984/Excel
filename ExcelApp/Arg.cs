// Reading arguments out of an interface request (moved out of Bridge.cs so code without the window can use it).
using System.Text.Json;

namespace ExcelApp;

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
