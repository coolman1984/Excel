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
    internal static FileStream Acquire(FileRoot root, string name) {
        string path = root.PathOf(name + ".lock"); Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var until = DateTime.UtcNow.AddSeconds(5);
        while (true) { try { return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) when (DateTime.UtcNow < until) { Thread.Sleep(50); } }
    }
    public CatalogData Snapshot() { lock (gate) return Read(); }
    public int NextVersion() => Update(c => checked(c.NextVersionId++));
}
