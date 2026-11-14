// Plan 4, lead: the editor's auto-save journal and crash recovery (phase 0).
// After every change the open workbook's edits are written, debounced, to the store as journal/<key>.json
// (cell data, so only through StoreRoot). Saving or discarding removes it. Opening a file that still has a
// journal offers the unsaved edits back. Nothing here reads the company file; the journal holds only the
// user's own edits, plus the file's length and time to tell whether the file changed since.
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelApp.Data;

namespace ExcelApp;

public sealed class EditJournal : IDisposable
{
    public sealed record Cell(int S, int R, int C, CellEdit E);
    public sealed record Width(int S, int C, double Px);
    public sealed record Freeze(int S, int Rows, int Cols);
    /// <summary>Ops: inserted and deleted rows and columns, in order; the cells' positions are after them.</summary>
    public sealed record Snapshot(string File, long Length, DateTime WriteUtc, DateTime SavedAt, List<Cell> Cells, List<Width> Widths, List<Freeze> Freezes, List<StructOp>? Ops = null);

    static readonly JsonSerializerOptions Opt = new() { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, Converters = { new JsonStringEnumConverter() } };
    readonly FileRoot root; readonly object gate = new();
    readonly System.Threading.Timer timer; Func<Snapshot?>? pending;
    public int DelayMs { get; init; } = 1500;

    public EditJournal(FileRoot root) { this.root = root; timer = new System.Threading.Timer(_ => Flush(), null, Timeout.Infinite, Timeout.Infinite); }

    public static string KeyOf(string file) =>
        "journal/" + Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(Path.GetFullPath(file).ToUpperInvariant())))[..24] + ".json";

    /// <summary>Takes a snapshot of the edits now (under the caller's edit lock) and writes it after a short quiet time.</summary>
    public static Snapshot Take(string file, EditSet edits)
    {
        var fi = new FileInfo(file);
        return new Snapshot(Path.GetFullPath(file), fi.Exists ? fi.Length : -1, fi.Exists ? fi.LastWriteTimeUtc : DateTime.MinValue, DateTime.Now,
            edits.All.Select(kv => new Cell(kv.Key.Sheet, kv.Key.Row, kv.Key.Col, kv.Value)).ToList(),
            edits.ColWidths.Select(kv => new Width(kv.Key.Sheet, kv.Key.Col, kv.Value)).ToList(),
            edits.Freeze.Select(kv => new Freeze(kv.Key, kv.Value.Rows, kv.Value.Cols)).ToList(),
            edits.Ops.Count == 0 ? null : edits.Ops.ToList());
    }

    /// <summary>Schedules a write. The snapshot is taken when the timer fires, so a burst of edits costs one write.</summary>
    public void Schedule(Func<Snapshot?> take) { lock (gate) { pending = take; timer.Change(DelayMs, Timeout.Infinite); } }

    /// <summary>Writes the pending snapshot now (also used on window close). An empty edit set removes the journal.</summary>
    public void Flush()
    {
        Func<Snapshot?>? take; lock (gate) { take = pending; pending = null; }
        if (take == null) return;
        try
        {
            var s = take(); if (s == null) return;
            string key = KeyOf(s.File);
            if (s.Cells.Count == 0 && s.Widths.Count == 0 && s.Freezes.Count == 0 && (s.Ops?.Count ?? 0) == 0) { root.Delete(key); return; }
            root.WriteAllText(key, JsonSerializer.Serialize(s, Opt));
        }
        catch (Exception e) { Log.Write("journal write " + e.GetType().Name); }
    }

    /// <summary>Drops a pending write and deletes the journal of this file (after a save, or when the user discards).</summary>
    public void Forget(string file)
    {
        lock (gate) { pending = null; timer.Change(Timeout.Infinite, Timeout.Infinite); }
        try { root.Delete(KeyOf(file)); } catch (Exception e) { Log.Write("journal delete " + e.GetType().Name); }
    }

    /// <summary>The journal of this file, or null. A broken journal is removed and reported as none.</summary>
    public Snapshot? Find(string file)
    {
        string key = KeyOf(file);
        try
        {
            if (!root.Exists(key)) return null;
            var s = JsonSerializer.Deserialize<Snapshot>(root.ReadAllText(key), Opt);
            if (s == null || !string.Equals(s.File, Path.GetFullPath(file), StringComparison.OrdinalIgnoreCase)) return null;
            return s;
        }
        catch (Exception e) when (e is JsonException or NotSupportedException) { Log.Write("journal unreadable"); try { root.Delete(key); } catch (IOException) { } return null; }
        catch (IOException e) { Log.Write("journal read " + e.GetType().Name); return null; }
    }

    /// <summary>True when the file on disk is the one the edits were made on.</summary>
    public static bool SameFile(Snapshot s)
    {
        var fi = new FileInfo(s.File);
        return fi.Exists && fi.Length == s.Length && fi.LastWriteTimeUtc == s.WriteUtc;
    }

    /// <summary>Puts the journal's edits back: its inserts and deletes first (in order), then the cells, one undo step per sheet.
    /// Cells outside the sheets are skipped. <paramref name="size"/> gives a sheet's size in the fast copy.</summary>
    public static int Restore(Snapshot s, EditSet edits, Func<int, (int Rows, int Cols)> size)
    {
        int n = 0;
        foreach (var op in s.Ops ?? new())
        {
            var (pr, pc) = size(op.Sheet); if (pr < 0) continue;
            try { edits.Structure(op, pr, pc); n++; } catch (ArgumentOutOfRangeException) { }
        }
        foreach (var g in s.Cells.GroupBy(c => c.S).OrderBy(g => g.Key))
        {
            var (pr, pc) = size(g.Key); if (pr < 0) continue;
            int rows = edits.RowCount(g.Key, pr), cols = edits.ColCount(g.Key, pc);
            var list = g.Where(c => c.R >= 0 && c.R < rows && c.C >= 0 && c.C < cols).Select(c => (c.R, c.C, (CellEdit?)c.E)).ToList();
            n += edits.Apply(g.Key, list).Count;
        }
        foreach (var w in s.Widths) if (size(w.S).Rows >= 0) edits.SetColWidth(w.S, w.C, w.Px);
        foreach (var f in s.Freezes) if (size(f.S).Rows >= 0) edits.SetFreeze(f.S, f.Rows, f.Cols);
        return n;
    }

    public void Dispose() { Flush(); timer.Dispose(); }
}
