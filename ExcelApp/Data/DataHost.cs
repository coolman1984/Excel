// Plan 4 (lead-owned): creates the shared services and registers every lane's "data.*" commands.
using System.Text.Json;
using ExcelViewer.Core;

namespace ExcelApp.Data;

public sealed class DataHost : IDisposable
{
    public DataServices? Services { get; }
    /// <summary>Plain sentence when the data engine cannot run here (for example, not the approved laptop).</summary>
    public string? Disabled { get; }

    DataHost(DataServices? s, string? disabled) { Services = s; Disabled = disabled; }

    public static DataHost Create(Action<string, object?> emit, StoreRoot? store = null, ConfigRoot? config = null)
    {
        try
        {
            var s = new DataServices { Emit = emit, Store = store ?? StoreRoot.Production(), Config = config ?? ConfigRoot.Production() };
            return new DataHost(s, null);
        }
        catch (CachePolicyException e) { return new DataHost(null, e.Message); }
        catch (UnauthorizedAccessException) { return new DataHost(null, "The data store folder cannot be opened."); }
        catch (IOException) { return new DataHost(null, "The data store folder cannot be opened."); }
    }

    public void Register(Dictionary<string, Func<JsonElement, Task<object?>>> h)
    {
        h["data.status"] = _ => Task.FromResult<object?>(new { enabled = Services != null, reason = Disabled });
        if (Services == null) return;
        DataCommands.Register(h, Services);   // lane A: store, workspace, ingestion, history, quality, export
        QueryCommands.Register(h, Services);  // lane B: views, joins, links, drill, alerts, recipes, sandbox, library, dashboard, questions, formulas
        Services.Ingestion?.Start();
    }

    public void Dispose()
    {
        try { Services?.Ingestion?.Stop(); } catch (Exception e) { Log.Write("ingestion stop " + e.GetType().Name); }
        if (Services == null) return;
        for (int i = Services.Owned.Count - 1; i >= 0; i--)
            try { Services.Owned[i].Dispose(); } catch (Exception e) { Log.Write("dispose " + e.GetType().Name); }
        Services.Owned.Clear();
    }
}
