// Runs every ExcelApp test. "offline" skips the tests that start a hidden Excel.
using System.Runtime.InteropServices;

static class TestProgram
{
    [DllImport("kernel32.dll")] static extern uint SetErrorMode(uint mode);
    static int pass, fail;
    static readonly object Gate = new();

    static void Check(string name, bool ok, string detail)
    {
        lock (Gate) { if (ok) pass++; else { fail++; Console.WriteLine($"FAIL {name} {detail}"); } }
    }

    [STAThread]
    static int Main(string[] a)
    {
        SetErrorMode(0x8003); // no Windows error box, ever
        bool offline = a.Contains("offline");
        var only = a.FirstOrDefault(x => x.StartsWith("--suites="))?.Substring(9).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        Selected = only is { Length: > 0 } ? new HashSet<string>(only, StringComparer.OrdinalIgnoreCase) : null;
        try
        {
            Suite("host", () => HostTests.Run(Check, offline));
            Suite("edit", () => EditTests.Run(Check));
            Suite("filter", () => FilterTests.Run(Check));
            Suite("save", () => SaveTests.Run(Check, offline));
            Suite("style", () => StyleTests.Run(Check, offline));
            Suite("journal", () => JournalTests.Run(Check));
            Suite("calc", () => CalcTests.Run(Check));
            // plan 4: lane A owns StoreTests/IngestTests, lane B owns QueryTests/AutomationTests/FormulaTests
            Suite("store", () => StoreTests.Run(Check, offline));
            Suite("ingest", () => IngestTests.Run(Check, offline));
            Suite("query", () => QueryTests.Run(Check, offline));
            Suite("automation", () => AutomationTests.Run(Check, offline));
            Suite("formula", () => FormulaTests.Run(Check, offline));
        }
        catch (Exception e) { Console.WriteLine("ERROR " + e.GetType().Name); return 1; }
        Console.WriteLine($"ExcelApp tests: {pass} passed, {fail} failed{(offline ? " (offline)" : "")}");
        return fail == 0 ? 0 : 1;
    }

    static HashSet<string>? Selected;   // null = every suite; "--suites=store,ingest" = only those

    static void Suite(string name, Action run)
    {
        if (Selected != null && !Selected.Contains(name)) return;
        try { run(); }
        catch (Exception e) { Check(name + " suite ran to the end", false, e.GetType().Name); }
    }
}
