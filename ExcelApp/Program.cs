// Modes: (none) window; --build <file>; --save; --selftest <made-up file> <png folder>; --dev <url>;
// AI agent access (read-only): --mcp (MCP server on stdio), --query "<SELECT ...>" [--max N], --tables.
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelViewer.Core;

namespace ExcelApp;

static class Program
{
    [DllImport("kernel32.dll")] static extern uint SetErrorMode(uint mode);
    [DllImport("kernel32.dll")] static extern bool AttachConsole(int processId);

    [STAThread]
    static int Main(string[] a)
    {
        SetErrorMode(0x8003); // no Windows error box, ever
        if (a.Length >= 1 && a[0] is "--mcp" or "--query" or "--tables") return AgentMode(a);
        bool unattended = a.Length >= 1 && a[0] is "--build" or "--save" or "--selftest" or "--export";
        if (a.Length == 2 && a[0] == "--export") { try { return ExcelApp.Data.ExportChild.Run(a[1]); } catch (Exception e) { Console.WriteLine("ERROR " + e.GetType().Name); return 1; } } // plan 4, lane A
        try
        {
            if (unattended && ((a[0] == "--build" && a.Length != 2) || (a[0] == "--save" && a.Length != 1) || (a[0] == "--selftest" && a.Length != 3))) { Console.WriteLine("ERROR arguments"); return 2; }
            if (a.Length >= 2 && a[0] == "--build") return BuildMode(a[1]);
            if (a.Length >= 1 && a[0] == "--save")
            {
                using var o = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
                return SaveBack.Run(Console.In, o);
            }
            ApplicationConfiguration.Initialize();
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s, e) =>
            {
                Log.Write("ui thread error " + e.Exception.GetType().Name);
                if (unattended) { Console.WriteLine("ERROR " + e.Exception.GetType().Name); Environment.Exit(1); }
                MessageBox.Show("Something went wrong (" + e.Exception.GetType().Name + "). Your file is not changed.", "ExcelApp");
            };
            if (a.Length >= 3 && a[0] == "--selftest") return SelfTest(a[1], a[2]);
            string? dev = null, file = null;
            if (a.Length >= 2 && a[0] == "--dev") dev = a[1];
            else if (a.Length >= 1 && File.Exists(a[0])) file = a[0];
            var f = new MainForm(dev) { OpenOnStart = file };
            Application.Run(f);
            return 0;
        }
        catch (Exception e)
        {
            try { Log.Write("fatal " + e.GetType().Name); } catch (Exception) { }
            if (unattended) Console.WriteLine("ERROR " + e.GetType().Name);
            else try { MessageBox.Show("ExcelApp could not start (" + e.GetType().Name + ").", "ExcelApp"); } catch (Exception) { }
            return 1;
        }
    }

    /// <summary>Read-only access for an AI agent to the data store (owner's decision, 2026-11-14). Never opens a window or writes data.</summary>
    static int AgentMode(string[] a)
    {
        bool mcp = a[0] == "--mcp";
        if (!mcp) AttachConsole(-1);   // a typed command shows its answer in the console it came from
        var utf8 = new UTF8Encoding(false);
        using var o = new StreamWriter(Console.OpenStandardOutput(), utf8) { AutoFlush = true };
        try
        {
            if (mcp)
            {
                // The workbook tools work whenever the window is open; the store tools need the approved store.
                ExcelApp.Data.Agent.AgentData? store = null; string? problem = null;
                try { var r = ExcelApp.Data.StoreRoot.Production(); store = new ExcelApp.Data.Agent.AgentData(new ExcelApp.Data.TableProvider(r, new ExcelApp.Data.Catalog(r))); }
                catch (Exception e) when (e is CachePolicyException or UserError) { problem = e.Message; }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { problem = "The data store folder cannot be opened."; }
                new ExcelApp.Data.Agent.McpServer(store, new ExcelApp.Agent.AppClient(), problem).Run(new StreamReader(Console.OpenStandardInput(), utf8), o);
                return 0;
            }
            var root = ExcelApp.Data.StoreRoot.Production();
            var api = new ExcelApp.Data.Agent.AgentData(new ExcelApp.Data.TableProvider(root, new ExcelApp.Data.Catalog(root)));
            object answer;
            if (a[0] == "--tables") answer = api.Tables();
            else
            {
                if (a.Length < 2) { o.WriteLine("ERROR Usage: ExcelApp.exe --query \"SELECT ...\" [--max N]"); return 2; }
                int max = ExcelApp.Data.Agent.AgentData.DefaultRows;
                int at = Array.IndexOf(a, "--max");
                if (at > 0 && (at + 1 >= a.Length || !int.TryParse(a[at + 1], NumberStyles.None, CultureInfo.InvariantCulture, out max))) { o.WriteLine("ERROR --max needs a whole number."); return 2; }
                answer = api.Query(a[1], max);
            }
            o.WriteLine(System.Text.Json.JsonSerializer.Serialize(answer, ExcelApp.Data.Agent.AgentData.JsonOptions));
            return 0;
        }
        catch (Exception e) when (e is UserError or CachePolicyException)
        {
            if (mcp) Console.Error.WriteLine(e.Message); else o.WriteLine("ERROR " + e.Message);
            return 1;
        }
        catch (Exception e)
        {
            Log.Write("agent " + e.GetType().Name);
            if (mcp) Console.Error.WriteLine("ExcelApp agent access stopped (" + e.GetType().Name + ")."); else o.WriteLine("ERROR " + e.GetType().Name);
            return 1;
        }
    }

    static int BuildMode(string source)
    {
        using var o = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
        bool gone = false;
        new Thread(() => { try { while (Console.In.ReadLine() != null) { } } catch (IOException) { } gone = true; }) { IsBackground = true }.Start();
        var r = WorkbookReader.Build(source, o.WriteLine, () => gone);
        o.WriteLine(string.Create(CultureInfo.InvariantCulture, $"RESULT code={r.Code} stopped={r.Stopped} excelStart={r.ExcelStartSec:F1}s open={r.OpenSec:F1}s formats={r.FormatSec:F1}s read={r.ReadSec:F1}s check={r.VerifySec:F1}s sheets={r.Sheets} cells={r.Cells} cacheMB={r.CacheBytes / 1048576.0:F1} formatsTotal={r.Formats} formatsMatched={r.FormatsOk} formatsNotCarried={r.FormatsFailed} formatsNoCells={r.FormatsUnused} unreadFormatRuns={r.UnreadRuns} peakExcelMB={r.PeakExcelMB} peakCommit={r.PeakCommit:F1}%"));
        return r.Code;
    }

    // ---- self test: the real window, off screen, driven through window.__app. Made-up files only. ----
    static long WebViewMB(MainForm f)
    {
        long total = 0;
        try
        {
            foreach (var info in f.View.CoreWebView2.Environment.GetProcessInfos())
                try { using var p = Process.GetProcessById(info.ProcessId); total += p.PrivateMemorySize64; } catch (ArgumentException) { }
        }
        catch (Exception) { return -1; }
        return total / 1048576;
    }

    static int SelfTest(string source, string folder)
    {
        if (!Path.GetFileName(source).StartsWith("syn", StringComparison.OrdinalIgnoreCase)) { Console.WriteLine("REFUSED: the self test is for made-up files only"); return 2; }
        Directory.CreateDirectory(folder);
        double fromStart = 0; int code = 0;
        var form = new MainForm { Unattended = true }; form.PlaceOffScreen(); form.ClientSize = new Size(1280, 760);
        form.Shown += async (s, e) =>
        {
            try { code = await RunSelfTest(form, source, folder); }
            catch (Exception ex) { Console.WriteLine("ERROR " + ex.GetType().Name); code = 1; }
            form.Close();
        };
        Application.Run(form);
        _ = fromStart; return code;
    }

    static async Task<string> Js(MainForm f, string script)
    {
        string r = await f.View.CoreWebView2.ExecuteScriptAsync(script);
        return r.Trim('"');
    }

    static async Task<bool> WaitJs(MainForm f, string expr, string want, double seconds)
    {
        var sw = Stopwatch.StartNew();
        while (sw.Elapsed.TotalSeconds < seconds) { if (await Js(f, expr) == want) return true; await Task.Delay(50); }
        return false;
    }

    static async Task Shot(MainForm f, string png)
    {
        using var ms = new MemoryStream();
        await f.View.CoreWebView2.CapturePreviewAsync(Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat.Png, ms);
        File.WriteAllBytes(png, ms.ToArray());
    }

    static async Task<int> RunSelfTest(MainForm f, string source, string folder)
    {
        var proc = Process.GetCurrentProcess();
        if (!await f.Ready) { Console.WriteLine("ERROR the page did not load"); return 1; }
        if (!await WaitJs(f, "typeof window.__app", "object", 30)) { Console.WriteLine("ERROR the interface has no self-test hooks"); return 1; }
        if (!await WaitJs(f, "window.__app.state()", "empty", 30)) { Console.WriteLine("ERROR the interface did not reach the empty state"); return 1; }
        double pageMs = (DateTime.Now - proc.StartTime).TotalMilliseconds;
        string cp = CachePolicy.PathFor(source);
        if (!CacheData.HeaderMatches(cp, source)) Console.WriteLine("NOTE no fresh copy yet: this run includes Excel reading the file");
        var sw = Stopwatch.StartNew();
        f.Session.Open(source);
        if (!await WaitJs(f, "window.__app.state()", "ready", 400)) { Console.WriteLine("ERROR the file did not open"); return 1; }
        if (!await WaitJs(f, "String(window.__app.hasRows())", "true", 30)) { Console.WriteLine("ERROR the first rows did not arrive"); return 1; }
        await Task.Delay(32);
        double total = (DateTime.Now - proc.StartTime).TotalMilliseconds;
        Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"FIRST SCREEN {total:F0} ms from process start (page ready at {pageMs:F0} ms, open to screen {sw.Elapsed.TotalMilliseconds:F0} ms)"));
        await Task.Delay(300); await Shot(f, Path.Combine(folder, "open.png"));
        var wb = (dynamic)f.Session.WorkbookInfo(); int rows = f.Session.Cache!.Sheets.Max(x => x.Rows);

        var rnd = new Random(7); var samples = new List<double?>();
        for (int i = 0; i < 10; i++)
        {
            int row = rnd.Next(0, Math.Max(1, rows - 40));
            // Duration is not an identity: distinct paints can have exactly equal timings.
            // Capture the completion token before triggering even an already-cached jump.
            string before = await Js(f, "String(window.__app.paintToken)");
            var jump = Stopwatch.StartNew();
            await Js(f, $"window.__app.scrollTo({row}); 'x'");
            double? elapsed = null;
            string after = before;
            while (jump.Elapsed.TotalSeconds < 5)
            {
                after = await Js(f, "String(window.__app.paintToken)");
                if (after != before && jump.Elapsed.TotalSeconds < 5)
                {
                    elapsed = jump.Elapsed.TotalMilliseconds;
                    break;
                }
                await Task.Delay(2);
            }
            samples.Add(elapsed);
            string loaded = await Js(f, $"String(window.__app.hasRows({row}))");
            Console.WriteLine($"JUMP PROBE {i + 1} row {row} token {before}->{after} loaded {loaded}: {JumpSamples.Format(elapsed)}");
        }
        Console.WriteLine(JumpSamples.Summary(samples));
        proc.Refresh(); Console.WriteLine($"MEMORY host {proc.PrivateMemorySize64 / 1048576} MB");
        Console.WriteLine($"MEMORY webview {WebViewMB(f)} MB");

        await Js(f, "window.__app.scrollTo(0); 'x'"); await Task.Delay(300);
        await Js(f, "window.__app.run('freeze', 2, 1); window.__app.run('zoom', 150); 'x'"); await Task.Delay(500);
        await Shot(f, Path.Combine(folder, "view.png"));
        await Js(f, "window.__app.run('zoom', 100); 'x'");
        await Js(f, "window.__styleDone=false; window.__app.run('select',1,1,3,2); window.__app.run('edit',1,1,'0.125').then(()=>window.__app.run('style',[1,1,3,2],{bold:true,fill:'#FFFF00',align:1,numberFormat:'0.00%'})).then(()=>window.__styleDone=true); 'x'");
        if (!await WaitJs(f, "String(window.__styleDone)", "true", 30)) { Console.WriteLine("ERROR style action failed"); return 1; }
        await Task.Delay(300);
        await Shot(f, Path.Combine(folder, "styles.png"));
        // Host-side timings through the same bridge the interface uses.
        var seen = new System.Collections.Concurrent.ConcurrentQueue<string>();
        var prior = f.Bridge.Post; f.Bridge.Post = j => { seen.Enqueue(j); prior?.Invoke(j); };
        int id = 1000;
        async Task<System.Text.Json.JsonElement> Cmd(string cmd, string args = "")
        {
            string reply = await f.Bridge.Dispatch("{\"id\":" + (id++) + ",\"cmd\":\"" + cmd + "\"" + (args.Length > 0 ? "," + args : "") + "}");
            using var doc = System.Text.Json.JsonDocument.Parse(reply);
            if (!doc.RootElement.GetProperty("ok").GetBoolean()) throw new InvalidOperationException(cmd);
            return doc.RootElement.GetProperty("result").Clone();
        }
        async Task<bool> WaitEvent(string name, double secs)
        {
            var w = Stopwatch.StartNew();
            while (w.Elapsed.TotalSeconds < secs) { foreach (var j in seen) if (j.Contains("\"event\":\"" + name + "\"")) return true; await Task.Delay(20); }
            return false;
        }
        var detail = await Cmd("cellDetail", "\"sheet\":0,\"row\":1,\"col\":1");
        bool styleOk = detail.GetProperty("style").GetProperty("bold").GetBoolean() && detail.GetProperty("style").GetProperty("fill").GetString() == "#FFFF00" && detail.GetProperty("format").GetString() == "0.00%";
        Console.WriteLine("STYLE wire/detail: " + (styleOk ? "pass" : "fail"));
        if (!styleOk) return 1;
        var cache = f.Session.Cache!; int big = 0; for (int i = 0; i < cache.Sheets.Count; i++) if (cache.Sheets[i].Cells > cache.Sheets[big].Cells) big = i;
        await Cmd("sheet", $"\"sheet\":{big}");
        int numCol = -1, textCol = -1; { var cells = (await Cmd("rows", $"\"sheet\":{big},\"from\":1,\"count\":1")).GetProperty("rows")[0][1];
            for (int k = 0; k < cells.GetArrayLength(); k++) { int kind = cells[k][1].GetInt32(); if (kind == 1 && numCol < 0) numCol = k; if (kind == 2 && textCol < 0) textCol = k; } }
        if (textCol >= 0)
        {
            var fv = await Cmd("filterValues", $"\"sheet\":{big},\"col\":{textCol}");
            string first = fv.GetProperty("values")[0].GetProperty("text").GetString()!;
            string valuesJson = System.Text.Json.JsonSerializer.Serialize(new[] { first });
            await Js(f, $"window.__app.run('filter', {textCol}, {valuesJson}); 'x'"); await Task.Delay(1500);
            await Shot(f, Path.Combine(folder, "filter.png"));
            await Cmd("clearFilters", $"\"sheet\":{big}");
        }
        if (numCol >= 0)
        {
            var w = Stopwatch.StartNew();
            var r = await Cmd("setFilter", $"\"sheet\":{big},\"col\":{numCol},\"condition\":{{\"kind\":\"number\",\"op\":\"gt\",\"a\":5000}}");
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"FILTER one column: {w.Elapsed.TotalMilliseconds:F0} ms, {r.GetProperty("viewRows").GetInt32():N0} rows shown"));
            await Cmd("clearFilters", $"\"sheet\":{big}");
            w.Restart(); await Cmd("setSort", $"\"sheet\":{big},\"keys\":[{{\"col\":{numCol},\"desc\":false}}]");
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"SORT one column: {w.Elapsed.TotalMilliseconds:F0} ms"));
            await Cmd("setSort", $"\"sheet\":{big},\"keys\":[]");
        }
        { var w = Stopwatch.StartNew(); while (seen.TryDequeue(out _)) { }
          await Cmd("find", "\"sheet\":-1,\"query\":\"NO-SUCH-VALUE-ZZZ\"");
          bool ok = await WaitEvent("findDone", 120);
          Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"FIND nothing found, every sheet: {w.Elapsed.TotalSeconds:F2} s{(ok ? "" : " (timed out)")}")); }
        Console.WriteLine($"MEMORY host after work {Process.GetCurrentProcess().PrivateMemorySize64 / 1048576} MB");

        // A save round trip of 100 edits on the smaller made-up file, from the save command to the reopened workbook.
        string small = Path.Combine(Path.GetDirectoryName(source)!, "syn-2m.xlsx");
        if (File.Exists(small))
        {
            while (seen.TryDequeue(out _)) { }
            f.Session.Open(small, discard: true);
            if (!await WaitEvent("opened", 400)) { Console.WriteLine("ERROR the small file did not open"); return 1; }
            await Task.Delay(300); while (seen.TryDequeue(out _)) { }
            for (int i = 0; i < 100; i++) await Cmd("setCell", $"\"sheet\":0,\"row\":{i + 1},\"col\":0,\"input\":\"{i}\"");
            var w = Stopwatch.StartNew();
            await Cmd("save", "\"saveAs\":false");
            bool ok = await WaitEvent("opened", 300);
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"SAVE 100 edits on the small file, save to reopened: {w.Elapsed.TotalSeconds:F1} s{(ok ? "" : " (no reopen)")}"));
            Console.WriteLine("SAVE file kept NASCA mark: " + (File.ReadAllBytes(small).Take(8).SequenceEqual(Encoding.ASCII.GetBytes("<## NASC")) ? "yes" : "no"));
        }
        Console.WriteLine("SELFTEST done");
        return 0;
    }
}
