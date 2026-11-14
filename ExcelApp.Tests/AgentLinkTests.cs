// AI agents controlling the running app: Excel addresses, the named-pipe door and the MCP workbook tools. No window needed.
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelApp.Agent;
using ExcelApp.Data.Agent;

public static class AgentLinkTests
{
    sealed class FakeApp : IAppLink
    {
        public readonly List<(string Cmd, string Args)> Calls = new();
        public JsonElement Call(string cmd, JsonElement args)
        {
            Calls.Add((cmd, args.GetRawText()));
            if (cmd == "workbook") return JsonSerializer.SerializeToElement(new { name = "syn.xlsx", dirty = true, currentSheet = 0, sheets = new[] { new { index = 0, name = "Sheet1", hidden = false, rows = 9, cols = 3, firstRow = 2, firstCol = 1, colWidthsPx = new[] { 64, 64, 64 } } } });
            if (cmd == "nope") throw new ExcelApp.UserError("That command is not known.");
            return JsonSerializer.SerializeToElement(new { ok = cmd });
        }
    }

    public static void Run(Action<string, bool, string> check)
    {
        // ---- Excel addresses
        check("a1: one cell", A1.Parse("B7") == (7, 2, 7, 2), "");
        check("a1: a range in either order, with $", A1.Parse("$D$20:a1") == (1, 1, 20, 4), A1.Parse("$D$20:a1").ToString());
        check("a1: whole column and rows stop at the used size", A1.Parse("C:D", 500, 30) == (1, 3, 500, 4) && A1.Parse("3:5", 500, 30) == (3, 1, 5, 30), "");
        check("a1: three-letter columns", A1.Parse("XFD1048576") == (1048576, 16384, 1048576, 16384) && A1.ColLetters(16384) == "XFD" && A1.ColNumber("AA") == 27, "");
        string Bad(string t) { try { A1.Parse(t); return "accepted"; } catch (ExcelApp.UserError e) { return e.Message; } }
        check("a1: nonsense is a plain sentence", Bad("hello").Contains("B7") && Bad("A0") != "accepted" && Bad("XFE1") != "accepted" && Bad("C") != "accepted" && Bad("A1:3") != "accepted", Bad("hello"));
        check("a1: format", A1.Format(1, 1, 20, 4) == "A1:D20" && A1.Format(7, 2, 7, 2) == "B7", "");

        // ---- the named-pipe door: the MCP side reaches the window's command dispatcher
        string pipe = "ExcelApp.Test." + Guid.NewGuid().ToString("N")[..8];
        var changedCmds = new List<string>();
        using (var link = new AgentLink(json =>
        {
            using var d = JsonDocument.Parse(json); var cmd = d.RootElement.GetProperty("cmd").GetString();
            if (cmd == "boom") return Task.FromResult("{\"id\":1,\"ok\":false,\"error\":\"Please wait until the current work is finished.\"}");
            return Task.FromResult(JsonSerializer.Serialize(new { id = 1, ok = true, result = new { cmd, sheet = d.RootElement.TryGetProperty("sheet", out var s) ? s.GetString() : null } }));
        }, cmd => { lock (changedCmds) changedCmds.Add(cmd); }, pipe))
        {
            link.Start();
            var client = new AppClient(pipe, 3000);
            var r = client.Call("agent.write", JsonSerializer.SerializeToElement(new { sheet = "Data", start = "B2" }));
            check("link: a command reaches the app and answers", r.GetProperty("cmd").GetString() == "agent.write" && r.GetProperty("sheet").GetString() == "Data", r.GetRawText());
            client.Call("agent.read", JsonSerializer.SerializeToElement(new { range = "A1" }));
            string err; try { client.Call("boom", JsonSerializer.SerializeToElement(new { })); err = "none"; } catch (ExcelApp.UserError e) { err = e.Message; }
            check("link: the app's refusal comes back as its sentence", err.Contains("Please wait"), err);
            var many = Enumerable.Range(0, 8).Select(i => Task.Run(() => client.Call("setStyle", JsonSerializer.SerializeToElement(new { sheet = "S" + i })).GetProperty("sheet").GetString())).ToArray();
            Task.WaitAll(many);
            check("link: several agents at once", many.Select(t => t.Result).OrderBy(x => x).SequenceEqual(Enumerable.Range(0, 8).Select(i => "S" + i).OrderBy(x => x)), "");
            SpinWait.SpinUntil(() => { lock (changedCmds) return changedCmds.Count >= 9; }, 2000);
            lock (changedCmds) check("link: the window redraws after changes, not after reads or refusals", changedCmds.Count == 9 && changedCmds[0] == "agent.write" && !changedCmds.Contains("agent.read") && !changedCmds.Contains("boom"), string.Join(",", changedCmds));
        }
        string closed; try { new AppClient(pipe + "x", 200).Call("workbook", JsonSerializer.SerializeToElement(new { })); closed = "none"; } catch (ExcelApp.UserError e) { closed = e.Message; }
        check("link: with no window open the agent is told plainly", closed.Contains("ExcelApp is not open"), closed);

        // ---- MCP workbook tools, through a stand-in for the window
        var app = new FakeApp(); var mcp = new McpServer(null, app, "The data store is only on the approved laptop.");
        JsonElement Call(string tool, object args) => JsonDocument.Parse(mcp.Handle(JsonSerializer.Serialize(new { jsonrpc = "2.0", id = 7, method = "tools/call", @params = new { name = tool, arguments = args } }))!).RootElement.GetProperty("result");
        var tools = JsonDocument.Parse(mcp.Handle("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/list\"}")!).RootElement.GetProperty("result").GetProperty("tools").EnumerateArray().Select(t => t.GetProperty("name").GetString()).ToArray();
        check("mcp app: workbook tools are offered", tools.SequenceEqual(new[] { "list_tables", "describe_table", "query", "app_status", "read_range", "write_range", "run_command" }), string.Join(",", tools));
        var status = Call("app_status", new { });
        var text = status.GetProperty("content")[0].GetProperty("text").GetString()!;
        check("mcp app: status gives sheets and used ranges, not column widths", text.Contains("\"used\":\"A2:C10\"") && text.Contains("\"Sheet1\"") && !text.Contains("colWidths"), text);
        Call("read_range", new { range = "'My Sheet'!B2:C3" });
        check("mcp app: Sheet!range is split into sheet and range", app.Calls[^1].Cmd == "agent.read" && app.Calls[^1].Args.Contains("\"sheet\":\"My Sheet\"") && app.Calls[^1].Args.Contains("\"range\":\"B2:C3\""), app.Calls[^1].Args);
        Call("write_range", new { start = "D4", sheet = "Data", values = new object[][] { new object[] { "x", 1, "=A1*2" } } });
        check("mcp app: write_range passes rows of values", app.Calls[^1].Cmd == "agent.write" && app.Calls[^1].Args.Contains("\"start\":\"D4\"") && app.Calls[^1].Args.Contains("\"=A1*2\"") && app.Calls[^1].Args.Contains("\"sheet\":\"Data\""), app.Calls[^1].Args);
        Call("run_command", new { command = "setSort", args = new { sheet = 0, keys = new[] { new { col = 1, desc = true } } } });
        check("mcp app: run_command passes any command", app.Calls[^1].Cmd == "setSort" && app.Calls[^1].Args.Contains("\"desc\":true"), app.Calls[^1].Args);
        var refused = Call("run_command", new { command = "nope" });
        check("mcp app: a refused command is a tool error", refused.GetProperty("isError").GetBoolean() && refused.GetProperty("content")[0].GetProperty("text").GetString()!.Contains("not known"), refused.GetRawText());
        var noStore = Call("query", new { sql = "SELECT 1" });
        check("mcp app: store tools explain why the store is missing", noStore.GetProperty("isError").GetBoolean() && noStore.GetProperty("content")[0].GetProperty("text").GetString()!.Contains("approved laptop"), noStore.GetRawText());
    }
}
