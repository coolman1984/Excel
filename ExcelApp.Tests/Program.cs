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
        try
        {
            Suite("host", () => HostTests.Run(Check, offline));
            Suite("edit", () => EditTests.Run(Check));
            Suite("filter", () => FilterTests.Run(Check));
            Suite("save", () => SaveTests.Run(Check, offline));
            Suite("style", () => StyleTests.Run(Check, offline));
        }
        catch (Exception e) { Console.WriteLine("ERROR " + e.GetType().Name); return 1; }
        Console.WriteLine($"ExcelApp tests: {pass} passed, {fail} failed{(offline ? " (offline)" : "")}");
        return fail == 0 ? 0 : 1;
    }

    static void Suite(string name, Action run)
    {
        try { run(); }
        catch (Exception e) { Check(name + " suite ran to the end", false, e.GetType().Name); }
    }
}
