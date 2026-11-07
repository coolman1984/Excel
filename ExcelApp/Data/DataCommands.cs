// Plan 4, LANE A owns this file (stub written by the lead so the base builds).
using System.Text.Json;

namespace ExcelApp.Data;

public static class DataCommands
{
    /// <summary>Creates lane A's services (sets s.Tables, s.Ingestion, s.Exporter) and registers the lane A "data.*" commands of plan 4 section 4.</summary>
    public static void Register(Dictionary<string, Func<JsonElement, Task<object?>>> h, DataServices s)
    {
    }
}
