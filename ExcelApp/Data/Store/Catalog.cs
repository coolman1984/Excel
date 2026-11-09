using System.Text.Json;
namespace ExcelApp.Data;

public sealed record WorkspaceEntry(TableInfo Info, string RelativePath, SourceInfo Source, long SourceLength = 0, long SourceTicks = 0);
public sealed class CatalogData
{
    public List<WorkspaceEntry> Workspace { get; set; } = new();
    public int NextVersionId { get; set; } = 1;
    public List<TableInfo> Tables { get; set; } = new();
    public List<string> ReportTypes { get; set; } = new();
    public Dictionary<string, string> Projects { get; set; } = new();
}
public sealed class Catalog
{
    readonly StoreRoot root; readonly object gate;
    static readonly System.Collections.Concurrent.ConcurrentDictionary<string, object> Gates = new(StringComparer.OrdinalIgnoreCase);
    public Catalog(StoreRoot root) { this.root = root; gate = Gates.GetOrAdd(root.PathOf("catalog.json"), _ => new object()); }
    CatalogData Read() => root.Exists("catalog.json") ? JsonSerializer.Deserialize<CatalogData>(root.ReadAllText("catalog.json"), Json.Options) ?? throw new InvalidDataException("Invalid catalog.") : new();
    public T Update<T>(Func<CatalogData,T> edit) { lock (gate) { using var lease = Acquire(root, "catalog.json"); var data = Read(); T value = edit(data); root.WriteAllText("catalog.json", JsonSerializer.Serialize(data, Json.Options)); return value; } }
    internal static IDisposable Acquire(FileRoot root, string name)
    {
        // A kernel mutex coordinates host instances without creating files outside FileRoot.
        string path = root.PathOf(name).ToUpperInvariant();
        string hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(path)));
        var mutex = new Mutex(false, @"Local\ExcelAppData-" + hash);
        try
        {
            bool entered;
            try { entered = mutex.WaitOne(TimeSpan.FromSeconds(5)); }
            catch (AbandonedMutexException) { entered = true; }
            if (!entered) throw new UserError("The data store is busy. Try again shortly.");
            return new Lease(mutex);
        }
        catch { mutex.Dispose(); throw; }
    }
    sealed class Lease(Mutex mutex) : IDisposable
    {
        public void Dispose() { mutex.ReleaseMutex(); mutex.Dispose(); }
    }
    public CatalogData Snapshot() { lock (gate) return Read(); }
    public int NextVersion() => Update(c => checked(c.NextVersionId++));
}
