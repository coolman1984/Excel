// Plan 4 (lead-owned): the ONLY way any lane creates, replaces, reads or deletes a store or config file.
// Production data: D:\ExcelViewerCache\store\ (CachePolicy.ValidateStore). Tests: a made-up folder under D:\ExcelAppTest\syn*.
using ExcelViewer.Core;

namespace ExcelApp.Data;

public abstract class FileRoot
{
    public string Folder { get; }
    public bool IsTest { get; }
    protected FileRoot(string folder, bool isTest) { Folder = Path.GetFullPath(folder); IsTest = isTest; }

    protected abstract void Check(string fullPath);

    /// <summary>Full path of a relative path ("reports/out/2026-10/v3.xvt"). Refuses rooted paths and "..".</summary>
    public string PathOf(string relative)
    {
        if (string.IsNullOrWhiteSpace(relative) || Path.IsPathRooted(relative) || relative.Split('/', '\\').Any(p => p is ".." or "."))
            throw new UserError("That store path is not allowed.");
        string full = Path.GetFullPath(Path.Combine(Folder, relative));
        if (!full.StartsWith(Folder + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) throw new UserError("That store path is not allowed.");
        Check(full);
        return full;
    }

    /// <summary>Opens "relative.part" for writing. Call Commit(relative) after closing it, so readers never see half a file.</summary>
    public FileStream OpenWrite(string relative)
    {
        string full = PathOf(relative) + ".part"; Check(full);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        return new FileStream(full, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 20);
    }

    public void Commit(string relative)
    {
        string full = PathOf(relative);
        File.Move(full + ".part", full, true);
    }

    /// <summary>Writes the whole text atomically (part file, then move).</summary>
    public void WriteAllText(string relative, string text)
    {
        using (var s = OpenWrite(relative)) using (var w = new StreamWriter(s, new System.Text.UTF8Encoding(false))) w.Write(text);
        Commit(relative);
    }

    /// <summary>Appends one line (JSON lines logs). Creates the file when missing.</summary>
    public void AppendLine(string relative, string line)
    {
        string full = PathOf(relative); Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        using var s = new FileStream(full, FileMode.Append, FileAccess.Write, FileShare.Read);
        using var w = new StreamWriter(s, new System.Text.UTF8Encoding(false)); w.WriteLine(line);
    }

    public FileStream OpenRead(string relative) => new(PathOf(relative), FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20);
    public bool Exists(string relative) => File.Exists(PathOf(relative));
    public string ReadAllText(string relative) => File.ReadAllText(PathOf(relative));
    public void Delete(string relative) { string f = PathOf(relative); if (File.Exists(f)) File.Delete(f); }

    /// <summary>Relative paths ("/" separated) of files under a relative folder matching a pattern ("*.xvt").</summary>
    public IReadOnlyList<string> List(string relativeFolder, string pattern)
    {
        string dir = relativeFolder.Length == 0 ? Folder : Path.Combine(Folder, relativeFolder);
        if (!Directory.Exists(dir)) return Array.Empty<string>();
        return Directory.GetFiles(dir, pattern, SearchOption.AllDirectories)
            .Select(f => Path.GetRelativePath(Folder, f).Replace('\\', '/')).Where(f => !f.EndsWith(".part", StringComparison.OrdinalIgnoreCase)).OrderBy(f => f, StringComparer.Ordinal).ToList();
    }

    protected static string TestFolder(string folder)
    {
        string full = Path.GetFullPath(folder);
        if (!full.StartsWith(@"D:\ExcelAppTest\", StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(full.TrimEnd('\\')).StartsWith("syn", StringComparison.OrdinalIgnoreCase))
            throw new UserError("Test folders must be under D:\\ExcelAppTest and start with 'syn'.");
        return full;
    }
}

/// <summary>Cell data: tables, versions, ingestion log, sandboxes, export jobs, journals.</summary>
public sealed class StoreRoot : FileRoot
{
    StoreRoot(string folder, bool isTest) : base(folder, isTest) { }
    public static StoreRoot Production() { CachePolicy.EnsureFolderLocked(); Directory.CreateDirectory(CachePolicy.StoreFolder); return new StoreRoot(CachePolicy.StoreFolder, false); }
    public static StoreRoot ForTest(string folder) { var f = TestFolder(folder); Directory.CreateDirectory(f); return new StoreRoot(f, true); }
    protected override void Check(string fullPath) { if (!IsTest) CachePolicy.ValidateStore(fullPath); }
}

/// <summary>A library file (.exlib) at a path the user chose (save or open dialog). Configuration only: the library test
/// proves no table data is inside. The only lane-facing way to touch a file outside the two roots.</summary>
public static class ExternalLibrary
{
    static string Checked(string fullPath)
    {
        if (string.IsNullOrWhiteSpace(fullPath) || !Path.IsPathRooted(fullPath)) throw new UserError("Choose a full path for the library file.");
        string f = Path.GetFullPath(fullPath);
        if (!f.EndsWith(".exlib", StringComparison.OrdinalIgnoreCase)) throw new UserError("A library file must end with .exlib.");
        if (f.StartsWith(CachePolicy.StoreFolder, StringComparison.OrdinalIgnoreCase)) throw new UserError("A library file cannot be saved in the data store.");
        return f;
    }
    /// <summary>Writes atomically (part file, then move).</summary>
    public static void Write(string fullPath, string text)
    {
        string f = Checked(fullPath); Directory.CreateDirectory(Path.GetDirectoryName(f)!);
        File.WriteAllText(f + ".part", text, new System.Text.UTF8Encoding(false)); File.Move(f + ".part", f, true);
    }
    public static string Read(string fullPath)
    {
        string f = Checked(fullPath);
        if (!File.Exists(f)) throw new UserError("That library file was not found.");
        if (new FileInfo(f).Length > 20L << 20) throw new UserError("That library file is too large.");
        return File.ReadAllText(f);
    }
}

/// <summary>Configuration only (recipes, views, dashboards, alerts, schedules, settings): %APPDATA%\ExcelApp\config.</summary>
public sealed class ConfigRoot : FileRoot
{
    ConfigRoot(string folder, bool isTest) : base(folder, isTest) { }
    public static ConfigRoot Production()
    {
        string f = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelApp", "config");
        Directory.CreateDirectory(f); return new ConfigRoot(f, false);
    }
    public static ConfigRoot ForTest(string folder) { var f = TestFolder(folder); Directory.CreateDirectory(f); return new ConfigRoot(f, true); }
    protected override void Check(string fullPath)
    {
        if (!fullPath.EndsWith(".json", StringComparison.OrdinalIgnoreCase) && !fullPath.EndsWith(".json.part", StringComparison.OrdinalIgnoreCase) &&
            !fullPath.EndsWith(".exlib", StringComparison.OrdinalIgnoreCase) && !fullPath.EndsWith(".exlib.part", StringComparison.OrdinalIgnoreCase))
            throw new UserError("Only configuration files may be written to the configuration folder.");
    }
}
