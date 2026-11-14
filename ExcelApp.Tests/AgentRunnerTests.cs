// The Agent panel's engine: Claude Code / Codex event parsing, command lines, and a full run against a stand-in CLI.
using System.Text.Json;
using ExcelApp.Agent;
using ExcelApp.Data;

public static class AgentRunnerTests
{
    public static void Run(Action<string, bool, string> check)
    {
        string? sid;
        // ---- Claude Code stream-json (shapes recorded from claude 2.1)
        var init = AgentRunner.ParseClaude("{\"type\":\"system\",\"subtype\":\"init\",\"session_id\":\"s-1\",\"mcp_servers\":[{\"name\":\"excelapp\",\"status\":\"connected\"}]}", out sid).ToList();
        check("claude: init gives the session and no noise", sid == "s-1" && init.Count == 0, string.Join("|", init));
        var bad = AgentRunner.ParseClaude("{\"type\":\"system\",\"subtype\":\"init\",\"session_id\":\"s-1\",\"mcp_servers\":[{\"name\":\"excelapp\",\"status\":\"failed\"}]}", out sid).ToList();
        check("claude: tools that failed to start are reported", bad.Count == 1 && bad[0].Kind == "error" && bad[0].Text.Contains("failed"), string.Join("|", bad));
        var said = AgentRunner.ParseClaude("{\"type\":\"assistant\",\"message\":{\"content\":[{\"type\":\"thinking\",\"thinking\":\"hm\"},{\"type\":\"text\",\"text\":\"Total is 42\"},{\"type\":\"tool_use\",\"name\":\"mcp__excelapp__query\",\"input\":{\"sql\":\"SELECT SUM(x) FROM t\"}}]}}", out sid).ToList();
        check("claude: text and tool calls, thinking left out", said.Count == 2 && said[0] == new AgentEvent("text", "Total is 42") && said[1] == new AgentEvent("tool", "query: SELECT SUM(x) FROM t"), string.Join("|", said));
        var failedTool = AgentRunner.ParseClaude("{\"type\":\"user\",\"message\":{\"content\":[{\"type\":\"tool_result\",\"is_error\":true,\"content\":[{\"type\":\"text\",\"text\":\"There is no column named X\"}]}]}}", out sid).ToList();
        check("claude: a failed tool shows its sentence", failedTool.Count == 1 && failedTool[0].Kind == "toolError" && failedTool[0].Text.Contains("no column"), string.Join("|", failedTool));
        var result = AgentRunner.ParseClaude("{\"type\":\"result\",\"subtype\":\"success\",\"is_error\":false,\"result\":\"Total is 42\",\"duration_ms\":2520,\"session_id\":\"s-2\"}", out sid).ToList();
        check("claude: result ends the turn and updates the session", sid == "s-2" && result.Single() == new AgentEvent("done", "Done in 2.5 s."), string.Join("|", result));
        var errResult = AgentRunner.ParseClaude("{\"type\":\"result\",\"subtype\":\"error_during_execution\",\"is_error\":true,\"result\":\"Not logged in\",\"session_id\":\"s-2\"}", out sid).ToList();
        check("claude: an error result is an error", errResult.Single().Kind == "error" && errResult[0].Text.Contains("Not logged in"), string.Join("|", errResult));
        check("claude: stream deltas and other events are ignored", !AgentRunner.ParseClaude("{\"type\":\"stream_event\",\"event\":{\"type\":\"content_block_delta\"}}", out sid).Any() && !AgentRunner.ParseClaude("{\"type\":\"rate_limit_event\"}", out sid).Any(), "");

        // ---- Codex exec --json (shapes recorded from codex-cli 0.169)
        AgentRunner.ParseCodex("{\"type\":\"thread.started\",\"thread_id\":\"01a1d81e\"}", out sid).ToList();
        check("codex: thread.started gives the session", sid == "01a1d81e", sid ?? "null");
        check("codex: reconnecting is a status, not a failure", AgentRunner.ParseCodex("{\"type\":\"error\",\"message\":\"Reconnecting... 2/5 (unexpected status 401)\"}", out sid).Single().Kind == "status", "");
        var cx = AgentRunner.ParseCodex("{\"type\":\"item.started\",\"item\":{\"id\":\"i1\",\"type\":\"mcp_tool_call\",\"server\":\"excelapp\",\"tool\":\"read_range\",\"arguments\":{\"range\":\"A1:C3\"},\"status\":\"in_progress\"}}", out sid)
            .Concat(AgentRunner.ParseCodex("{\"type\":\"item.completed\",\"item\":{\"id\":\"i2\",\"type\":\"agent_message\",\"text\":\"Done: 3 rows\"}}", out sid))
            .Concat(AgentRunner.ParseCodex("{\"type\":\"turn.completed\",\"usage\":{\"input_tokens\":5}}", out sid)).ToList();
        check("codex: tool call, message and end of turn", cx.Count == 3 && cx[0] == new AgentEvent("tool", "read_range: A1:C3") && cx[1] == new AgentEvent("text", "Done: 3 rows") && cx[2].Kind == "done", string.Join("|", cx));
        check("codex: a failed turn is an error", AgentRunner.ParseCodex("{\"type\":\"turn.failed\",\"error\":{\"message\":\"401 Unauthorized\"}}", out sid).Single() == new AgentEvent("error", "401 Unauthorized"), "");
        check("codex: plain text lines are shown as status", AgentRunner.ParseCodex("WARNING: something", out sid).Single().Kind == "status", "");

        // ---- command lines: the request is never on the command line
        var (f1, a1) = AgentRunner.Command("claude", "claude", "m.json", "w", "ExcelApp.exe", "s-9");
        check("command: claude headless with ExcelApp's MCP server only", string.Join(" ", a1) == "-p --output-format stream-json --verbose --mcp-config m.json --strict-mcp-config --allowedTools mcp__excelapp --resume s-9", string.Join(" ", a1));
        var (f2, a2) = AgentRunner.Command("codex", "codex", "m.json", "w", "C:\\x\\ExcelApp.exe", "t-1");
        check("command: codex exec with the MCP server, resume, request on stdin", string.Join(" ", a2) == "exec --json --skip-git-repo-check -C w -c mcp_servers.excelapp.command='C:\\x\\ExcelApp.exe' -c mcp_servers.excelapp.args=['--mcp'] -c approval_policy='never' resume t-1 -", string.Join(" ", a2));

        // ---- a full run against a stand-in CLI that answers like Claude Code
        var config = ConfigRoot.ForTest(@"D:\ExcelAppTest\syn-agent-" + Guid.NewGuid().ToString("N")[..8]);
        string fake = MakeFakeClaude(StoreRoot.ForTest(@"D:\ExcelAppTest\syn-fakecli-" + Guid.NewGuid().ToString("N")[..8]));
        var events = new List<AgentEvent>(); var gate = new object();
        using var runner = new AgentRunner("ExcelApp.exe", config, e => { lock (gate) events.Add(e); }, name => name == "claude" ? fake : null);
        var st = JsonSerializer.SerializeToElement(runner.Status());
        check("run: status says which CLIs are installed", st.GetProperty("claude").GetBoolean() && !st.GetProperty("codex").GetBoolean(), st.GetRawText());
        string missing; try { runner.Start("codex", "hi", false); missing = "started"; } catch (ExcelApp.UserError e) { missing = e.Message; }
        check("run: a missing CLI says how to install it", missing.Contains("npm install -g @openai/codex"), missing);
        bool Wait(Func<List<AgentEvent>, bool> f) => SpinWait.SpinUntil(() => { lock (gate) return f(events); }, 15000);
        runner.Start("claude", "hello", false);
        bool first = Wait(e => e.Any(x => x.Kind is "done" or "error"));
        List<AgentEvent> got; lock (gate) got = events.ToList();
        check("run: the request reaches the CLI on its input, events come back", first && got.Any(e => e.Kind == "text" && e.Text.Contains("you said hello")) && got.Any(e => e == new AgentEvent("tool", "read_range: A1:B2")) && got.Last().Kind == "done", string.Join(" | ", got));
        check("run: the CLI gets ExcelApp's MCP config", got.Any(e => e.Kind == "text" && e.Text.Contains("--mcp-config") && e.Text.Contains("mcp.json")) && config.Exists("agent/mcp.json") && config.ReadAllText("agent/mcp.json").Contains("--mcp"), string.Join(" | ", got));
        SpinWait.SpinUntil(() => JsonSerializer.SerializeToElement(runner.Status()).GetProperty("running").ValueKind == JsonValueKind.Null, 5000);
        lock (gate) events.Clear();
        runner.Start("claude", "again", false);
        Wait(e => e.Any(x => x.Kind is "done" or "error"));
        lock (gate) got = events.ToList();
        check("run: the next message continues the same conversation", got.Any(e => e.Kind == "text" && e.Text.Contains("--resume fake-session")), string.Join(" | ", got));
        SpinWait.SpinUntil(() => JsonSerializer.SerializeToElement(runner.Status()).GetProperty("running").ValueKind == JsonValueKind.Null, 5000);
        lock (gate) events.Clear();
        runner.Start("claude", "slow", true);
        string busy; try { runner.Start("claude", "more", false); busy = "started"; } catch (ExcelApp.UserError e) { busy = e.Message; }
        check("run: one turn at a time", busy.Contains("still working"), busy);
        Thread.Sleep(500); var sw = System.Diagnostics.Stopwatch.StartNew(); runner.Stop();
        bool stopped = Wait(e => e.Any(x => x.Kind == "done" && x.Text == "Stopped."));
        lock (gate) got = events.ToList();
        check("run: Stop ends the agent at once", stopped && sw.ElapsedMilliseconds < 5000 && !got.Any(e => e.Text.Contains("--resume")), string.Join(" | ", got));
    }

    /// <summary>A stand-in for the claude CLI: echoes its request and arguments in Claude Code's JSON. "slow" waits 30 s.</summary>
    static string MakeFakeClaude(StoreRoot config)
    {
        if (OperatingSystem.IsWindows())
        {
            config.WriteAllText("fake/claude.cmd", string.Join("\r\n", new[] {
                "@echo off", "setlocal enabledelayedexpansion", "for /f \"delims=\" %%L in ('more') do set REQ=%%L", "set ARGS=%*", "set ARGS=!ARGS:\"=!", "set ARGS=!ARGS:\\=/!",
                "if \"%REQ%\"==\"slow\" ping -n 30 127.0.0.1 >nul",
                "echo {\"type\":\"system\",\"subtype\":\"init\",\"session_id\":\"fake-session\",\"mcp_servers\":[{\"name\":\"excelapp\",\"status\":\"connected\"}]}",
                "echo {\"type\":\"assistant\",\"message\":{\"content\":[{\"type\":\"tool_use\",\"name\":\"mcp__excelapp__read_range\",\"input\":{\"range\":\"A1:B2\"}}]}}",
                "echo {\"type\":\"assistant\",\"message\":{\"content\":[{\"type\":\"text\",\"text\":\"you said !REQ! with !ARGS!\"}]}}",
                "echo {\"type\":\"result\",\"subtype\":\"success\",\"is_error\":false,\"result\":\"ok\",\"duration_ms\":10,\"session_id\":\"fake-session\"}" }) + "\r\n");
            return config.PathOf("fake/claude.cmd");
        }
        config.WriteAllText("fake/claude", string.Join("\n", new[] {
            "#!/bin/sh", "REQ=$(cat | tail -n 1)", "ARGS=$(echo \"$*\" | tr -d '\"')",
            "if [ \"$REQ\" = slow ]; then sleep 30; fi",
            "echo '{\"type\":\"system\",\"subtype\":\"init\",\"session_id\":\"fake-session\",\"mcp_servers\":[{\"name\":\"excelapp\",\"status\":\"connected\"}]}'",
            "echo '{\"type\":\"assistant\",\"message\":{\"content\":[{\"type\":\"tool_use\",\"name\":\"mcp__excelapp__read_range\",\"input\":{\"range\":\"A1:B2\"}}]}}'",
            "echo \"{\\\"type\\\":\\\"assistant\\\",\\\"message\\\":{\\\"content\\\":[{\\\"type\\\":\\\"text\\\",\\\"text\\\":\\\"you said $REQ with $ARGS\\\"}]}}\"",
            "echo '{\"type\":\"result\",\"subtype\":\"success\",\"is_error\":false,\"result\":\"ok\",\"duration_ms\":10,\"session_id\":\"fake-session\"}'" }) + "\n");
        string path = config.PathOf("fake/claude");
        File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        return path;
    }
}
