// Plan 4 shared contract (lead-owned). Every lane builds on these types; no lane edits this file.
// If a lane needs a change here, it writes the request in its answer file and the lead decides.
//
// Rules that bind every type here (plan 4 section 0):
// - Cell data lives only under StoreRoot (D:\ExcelViewerCache\store on the approved laptop).
// - Configuration (recipes, views, dashboards, alerts) lives under ConfigRoot and holds no cell data
//   except literal rule values the user typed; the library export shows those before sharing.
// - Never log cell values, sheet names or source file names to Log.cs or the console.
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelApp.Data;

// =====================================================================================
// 1. Columns and tables (in memory). Row indexes are 0-based everywhere.
// =====================================================================================

public enum ColType : byte { Number = 1, Text = 2, Date = 3, Bool = 4 }

/// <summary>A read-only column. Dates are Excel serial days (1900 system, 1899-12-30 = 0).</summary>
public abstract class Column
{
    public string Name { get; }
    public ColType Type { get; }
    public int Rows { get; }
    protected Column(string name, ColType type, int rows) { Name = name; Type = type; Rows = rows; }
    public abstract bool IsNull(int row);
    /// <summary>Number, Date (serial days) or Bool (0/1). NaN when null or Text.</summary>
    public virtual double Num(int row) => double.NaN;
    /// <summary>Invariant display text: numbers "R" format, dates "yyyy-MM-dd" (with " HH:mm:ss" when there is a time part), bools TRUE/FALSE, null "".</summary>
    public abstract string Text(int row);
}

public sealed class NumberColumn : Column
{
    public readonly double[] Values; public readonly bool[]? Nulls; // Nulls == null: no nulls
    public NumberColumn(string name, double[] values, bool[]? nulls, ColType type = ColType.Number)
        : base(name, type, values.Length)
    {
        if (type is not (ColType.Number or ColType.Date)) throw new ArgumentException("NumberColumn holds Number or Date.");
        if (nulls != null && nulls.Length != values.Length) throw new ArgumentException("Nulls length differs.");
        Values = values; Nulls = nulls;
    }
    public override bool IsNull(int row) => Nulls != null && Nulls[row];
    public override double Num(int row) => IsNull(row) ? double.NaN : Values[row];
    public override string Text(int row)
    {
        if (IsNull(row)) return "";
        double v = Values[row];
        if (Type == ColType.Number) return v.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        var d = DateTime.FromOADate(v);
        return d.TimeOfDay == TimeSpan.Zero ? d.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
                                            : d.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
    }
}

/// <summary>Dictionary-encoded text. Codes[row] == -1 means null. Dict holds each distinct string once.</summary>
public sealed class TextColumn : Column
{
    public readonly int[] Codes; public readonly string[] Dict;
    public TextColumn(string name, int[] codes, string[] dict) : base(name, ColType.Text, codes.Length) { Codes = codes; Dict = dict; }
    public override bool IsNull(int row) => Codes[row] < 0;
    public override string Text(int row) => Codes[row] < 0 ? "" : Dict[Codes[row]];
    /// <summary>Builds a column from plain strings (null = null). For tests and small results.</summary>
    public static TextColumn From(string name, IReadOnlyList<string?> values)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal); var dict = new List<string>(); var codes = new int[values.Count];
        for (int i = 0; i < values.Count; i++)
        {
            var s = values[i]; if (s == null) { codes[i] = -1; continue; }
            if (!map.TryGetValue(s, out int c)) { c = dict.Count; map[s] = c; dict.Add(s); }
            codes[i] = c;
        }
        return new TextColumn(name, codes, dict.ToArray());
    }
}

/// <summary>Values: 0 false, 1 true, -1 null.</summary>
public sealed class BoolColumn : Column
{
    public readonly sbyte[] Values;
    public BoolColumn(string name, sbyte[] values) : base(name, ColType.Bool, values.Length) { Values = values; }
    public override bool IsNull(int row) => Values[row] < 0;
    public override double Num(int row) => Values[row] < 0 ? double.NaN : Values[row];
    public override string Text(int row) => Values[row] < 0 ? "" : Values[row] == 1 ? "TRUE" : "FALSE";
}

/// <summary>
/// An immutable table. Reserved lineage columns (plan 4 section 2.6), present on every loaded table:
///   "_v"  Number: the version number the row came from
///   "_r"  Number: the 1-based row number in the source sheet
/// Names starting with "_" are hidden in the interface and never typed or cleaned. One exception: "_status" of a diff
/// result (added|removed|changed|same) IS sent by data.page and listed in ResultInfo.columns, for row colouring.
/// </summary>
public sealed class Table
{
    public string Name { get; }
    public IReadOnlyList<Column> Columns { get; }
    public int Rows { get; }
    public Table(string name, IReadOnlyList<Column> columns)
    {
        Name = name; Columns = columns; Rows = columns.Count == 0 ? 0 : columns[0].Rows;
        foreach (var c in columns) if (c.Rows != Rows) throw new ArgumentException("Every column must have the same number of rows.");
    }
    public int IndexOf(string name) { for (int i = 0; i < Columns.Count; i++) if (string.Equals(Columns[i].Name, name, StringComparison.OrdinalIgnoreCase)) return i; return -1; }
    public Column? this[string name] { get { int i = IndexOf(name); return i < 0 ? null : Columns[i]; } }
}

// =====================================================================================
// 2. References to data (wire and C#). Lane A resolves them; lane B and the interface use them.
// =====================================================================================

/// <summary>
/// Kind "report": the current data of a report type (Latest or Locked truth, plan 2.6), optionally as of a time or version.
/// Kind "table":  a workspace table (one sheet of one file in a project, plan idea 2).
/// Kind "version": exactly one stored version (Id = version id).
/// Kind "sandbox": a what-if sandbox (Id = sandbox id), lane B.
/// </summary>
public sealed record TableRef(string Kind, string Id, string? AsOf = null, int? Version = null, string? Truth = null);
// AsOf: ISO "yyyy-MM-dd" or "yyyy-MM-ddTHH:mm:ss" (local). Truth: "latest" | "locked" (null = the report type's default).

public sealed record ColumnInfo(string Name, ColType Type);
public sealed record TableInfo(string Kind, string Id, string Name, string Group, int Rows, IReadOnlyList<ColumnInfo> Columns, DateTime LoadedAt);
// Group: report type name, or project name for workspace tables.

/// <summary>Where the rows of one stored version came from (plan idea 13). "_v" in a table holds VersionId.</summary>
public sealed record SourceInfo(
    int VersionId,                                   // global, unique in the store, increasing
    string Kind,                                     // "report" | "table"
    string GroupId,                                  // report type id or project id
    string Period,                                   // "" for workspace tables
    int PeriodVersion,                               // 1, 2, 3 ... within the period (0 for workspace tables)
    string SourcePath, string Sheet, int HeaderRow, DateTime LoadedAt, string LoadId,
    IReadOnlyDictionary<string, int> SourceColumn);  // table column name -> 1-based sheet column

/// <summary>Lane A implements; lane B and DataHost use it. Thread-safe. Tables are cached and shared read-only.</summary>
public interface ITableProvider
{
    Table Get(TableRef r);                         // throws UserError (plain sentence) when it does not exist
    /// <summary>The same table with only these columns loaded (plus lineage), so a question over 300 columns reads only what it uses.
    /// Null means every column. A provider may return more columns than asked; never fewer that exist.</summary>
    Table Get(TableRef r, IReadOnlyCollection<string>? columns) => Get(r);
    /// <summary>Only these rows (row numbers of Get's table, in this order) of these columns (null = all). A wide answer reads its
    /// 300 columns for the rows it shows instead of for every row.</summary>
    Table GetRows(TableRef r, IReadOnlyCollection<string>? columns, int[] rows)
    { var t = Get(r, columns); return new Table(t.Name, t.Columns.Select(c => ExcelApp.Data.Query.Engine.Gather(c, rows)).ToArray()); }
    IReadOnlyList<TableInfo> List();
    SourceInfo? SourceOf(int versionId);
    /// <summary>Stores a table as a new version of a report type's period (used by sandbox "keep", plan idea 15). Returns the period version number.</summary>
    int SaveAsVersion(string reportTypeId, string period, Table data, string note);
}

/// <summary>Results any lane produces and the interface pages through with "data.page" (lane B formats pages). Lead-owned.</summary>
public sealed class ResultStore
{
    sealed record Entry(Table Table, ResultMeta Meta, long Bytes);
    readonly Dictionary<string, Entry> map = new(); readonly LinkedList<string> order = new(); readonly object gate = new();
    long bytes; int next;
    public long MaxBytes { get; set; } = 1L << 30;   // 1 GB of results, then the oldest are released
    public string Add(Table t, ResultMeta meta)
    {
        long b = EstimateBytes(t);
        lock (gate)
        {
            string id = "r" + (++next).ToString(System.Globalization.CultureInfo.InvariantCulture);
            map[id] = new Entry(t, meta, b); order.AddLast(id); bytes += b;
            while (bytes > MaxBytes && order.Count > 1) { var old = order.First!.Value; order.RemoveFirst(); bytes -= map[old].Bytes; map.Remove(old); }
            return id;
        }
    }
    public bool TryGet(string id, out Table table, out ResultMeta meta)
    {
        lock (gate)
        {
            if (map.TryGetValue(id, out var e)) { order.Remove(id); order.AddLast(id); table = e.Table; meta = e.Meta; return true; }
            table = null!; meta = null!; return false;
        }
    }
    public void Release(string id) { lock (gate) if (map.Remove(id, out var e)) { order.Remove(id); bytes -= e.Bytes; } }
    static long EstimateBytes(Table t) { long b = 0; foreach (var c in t.Columns) b += c switch { TextColumn x => x.Codes.Length * 4L + x.Dict.Sum(s => (long)s.Length * 2), BoolColumn x => x.Values.Length, NumberColumn x => x.Values.Length * 9L, _ => 0 }; return b; }
}

/// <summary>Kind: "query" | "diff" | "quality" | "drill" | "sandbox". Spec: the lane's own spec object (JSON-serialisable) or null.
/// For "diff" results the table has a "_status" Text column: "added" | "removed" | "changed" | "same".</summary>
public sealed record ResultMeta(string Kind, object? Spec, bool Grouped, TableRef? Source);

// =====================================================================================
// 3. Ingestion (plan section 2). Lane A owns behaviour; these are the shapes everyone sees.
// =====================================================================================

public enum Granularity { Day, Week, Month }

public sealed class IngestRecipe
{
    public string Id { get; set; } = "";              // slug, unique, [a-z0-9-]
    public string Name { get; set; } = "";
    public List<string> Folders { get; set; } = new(); // absolute folders to watch
    public List<string> Patterns { get; set; } = new(); // file globs, e.g. "OutputReport_*.xlsx"
    public List<string> HeaderSignature { get; set; } = new(); // expected header names, in order
    public int HeaderRow { get; set; } = 0;           // 1-based; 0 = detect
    public string Sheet { get; set; } = "";           // sheet name; "" = the first visible sheet
    public string DateColumn { get; set; } = "";
    public Granularity Granularity { get; set; } = Granularity.Month;
    public WeekRule Week { get; set; } = new();
    public List<string> KeyColumns { get; set; } = new();
    public List<TypeFix> TypeFixes { get; set; } = new();
    public double PartialThreshold { get; set; } = 0.30; // plan 2.5: stop when rows fall by more than this share
    public string Truth { get; set; } = "latest";      // "latest" | "locked"
    public LockRule Lock { get; set; } = new();
    public int KeepVersions { get; set; } = 0;         // 0 = keep every version (plan 2.2: nothing is silently lost)
    public List<string> DependentViews { get; set; } = new();
}

public sealed class WeekRule
{
    public DayOfWeek FirstDay { get; set; } = DayOfWeek.Monday;
    public string Numbering { get; set; } = "iso";    // "iso" | "fiscal"
    public string? FiscalYearStart { get; set; }       // "MM-dd" when Numbering is fiscal, week 1 starts on the FirstDay on or before it
}

public sealed class LockRule
{
    public int WorkingDay { get; set; } = 5;          // the period locks at the end of this working day of the next period
    public List<string> Holidays { get; set; } = new(); // "yyyy-MM-dd", not working days
}

public sealed class TypeFix
{
    public string Column { get; set; } = "";
    public string Fix { get; set; } = "";             // "date" | "number" | "text" | "padZeros" | "trim" | "upper"
    public int Width { get; set; } = 0;               // padZeros width
    public string? DateFormat { get; set; }            // for text dates, e.g. "dd.MM.yyyy"; null = detect
}

/// <summary>One load of one file. Persisted in the ingestion log (store) and sent to the interface.</summary>
public sealed class LoadRecord
{
    public string Id { get; set; } = "";              // "L" + yyyyMMddHHmmssfff
    public string RecipeId { get; set; } = "";
    public string SourcePath { get; set; } = "";      // full path; shown only in the app, never logged
    public string FileHash { get; set; } = "";        // SHA-256 of the file bytes (identity only, never parsed)
    public string ContentHash { get; set; } = "";     // SHA-256 over the loaded cell values (detects identical re-downloads with new bytes)
    public DateTime LoadedAt { get; set; }
    public List<string> Periods { get; set; } = new(); // "2026-10", "2026-W44", "2026-10-31"
    public int Rows { get; set; }
    public int Added { get; set; } public int Removed { get; set; } public int Changed { get; set; }
    public string Status { get; set; } = "";          // "loaded" | "duplicate" | "quarantined" | "rejected" | "rolledBack" | "failed"
    public string? Reason { get; set; }                // plain sentence for quarantine/failed
    public Dictionary<string, int> Versions { get; set; } = new(); // period -> version number created
    public CleanupReport? Cleanup { get; set; }
    public QualityCard? Quality { get; set; }
}

public sealed class CleanupReport
{
    public List<CleanupItem> Items { get; set; } = new();
}
public sealed record CleanupItem(string Column, string Fix, int Cells, string Example); // Example: "text '01.10.2026' -> date" (no real values in logs; shown in app only)

public sealed class QualityCard
{
    public List<QualityFlag> Flags { get; set; } = new();
}
public sealed record QualityFlag(string Kind, string Column, int Count, string Message, IReadOnlyList<int> SampleRows);
// Kind: "duplicateKey" | "blank" | "badDate" | "outOfRange" | "outlier" | "newCategory" | "futureDate" | "twoPeriods"

/// <summary>Lane A implements. Raised after every successful load (and rollback), on a background thread.</summary>
public interface IIngestion
{
    event Action<LoadRecord>? Loaded;
    IReadOnlyList<IngestRecipe> Recipes { get; }
    void SaveRecipe(IngestRecipe r);
    void DeleteRecipe(string id);
    void Start();                                     // watchers + catch-up scan; idempotent
    void Stop();
}

// =====================================================================================
// 4. Export through Excel (plan ideas 18 and 19). Lane A implements the child; lane B builds jobs.
// =====================================================================================

public sealed class ExportJob
{
    public string Target { get; set; } = "";          // the .xlsx path the user chose (Excel writes it, so NASCA protects it)
    public string? Template { get; set; }             // optional .xlsx template, copied by Excel then filled
    public List<ExportSheet> Sheets { get; set; } = new();
    public bool Pdf { get; set; }                      // only when Settings.AllowPdf is true (owner decision, plan 4 section 7)
}
public sealed class ExportSheet
{
    public string Sheet { get; set; } = "";           // created when missing
    public string StartCell { get; set; } = "A1";     // or a defined name in the template
    public bool Headers { get; set; } = true;
    public List<string> Columns { get; set; } = new();
    public List<string?[]> Rows { get; set; } = new(); // invariant text per cell; numbers as "R" strings, dates as yyyy-MM-dd
    public List<string> NumberFormats { get; set; } = new(); // per column, "" = keep template/General
}

public interface IExporter
{
    /// <summary>Writes through a hidden Excel child (ExcelApp.exe --export). Throws UserError with a plain sentence on failure.</summary>
    Task ExportAsync(ExportJob job, IProgress<string>? progress, CancellationToken token);
}

// =====================================================================================
// 5. Shared services: DataHost creates one, lanes fill it in Register (A first, then B).
// =====================================================================================

public sealed class DataServices
{
    public required Action<string, object?> Emit { get; init; }  // host event to the interface: Emit("data.loaded", payload)
    public required StoreRoot Store { get; init; }
    public required ConfigRoot Config { get; init; }
    public ITableProvider Tables { get; set; } = null!;          // lane A sets
    public IIngestion Ingestion { get; set; } = null!;           // lane A sets
    public IExporter Exporter { get; set; } = null!;             // lane A sets
    public ResultStore Results { get; } = new();
    public Settings Settings { get; set; } = new();
    /// <summary>Services with timers or subscriptions (alerts, scheduler, watchers) add themselves here; DataHost.Dispose
    /// disposes them in reverse order, after stopping ingestion.</summary>
    public List<IDisposable> Owned { get; } = new();
}

public sealed class Settings
{
    public bool AllowPdf { get; set; } = false;       // owner/security decision pending (plan 4 section 7)
    public bool StartWithWindows { get; set; } = false;
    public int MaxParallelExcel { get; set; } = 1;     // one hidden Excel at a time
}

public static class Json
{
    public static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };
}
