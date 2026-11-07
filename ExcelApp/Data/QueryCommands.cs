// Plan 4, LANE B owns this file (stub written by the lead so the base builds).
using System.Text.Json;

namespace ExcelApp.Data;

public static class QueryCommands
{
    /// <summary>Creates lane B's services and registers the lane B "data.*" commands of plan 4 section 4. Runs after DataCommands.Register.</summary>
    public static void Register(Dictionary<string, Func<JsonElement, Task<object?>>> h, DataServices s)
    {
    }
}
