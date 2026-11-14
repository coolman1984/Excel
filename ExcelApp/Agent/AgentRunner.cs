// The Agent panel's engine: runs Claude Code CLI or Codex CLI (whatever the owner has signed in to; no API key in ExcelApp),
// gives it ExcelApp's MCP server (ExcelApp.exe --mcp), writes the request to its input and turns its JSON events into
// "agent.event" messages for the interface. The agent then works on the open workbook and the data store through MCP.
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using ExcelApp.Data;

namespace ExcelApp.Agent;

/// <summary>One line for the Agent panel. Kind: session, text, tool, toolError, status, error, done.</summary>
public sealed record AgentEvent(string Kind, string Text);

public sealed class AgentRunner : IDisposable
{
    public const string Instructions =
        "You are working inside ExcelApp, an Excel-like program, for its owner. Use the excelapp tools: app_status, read_range, " +
        "write_range and run_command for the workbook open in the window; list_tables, describe_table and query (SQL) for the data store. " +
        "Do the work through these tools yourself; do not ask the owner to do it by hand. Changes stay unsaved until run_command save " +
        "or the owner saves. Answer in the owner's language, briefly.";

    readonly string exePath; readonly ConfigRoot config; readonly Action<AgentEvent> emit;
    readonly Func<string, string?> find; readonly object gate = new();
    Process? proc; string? running; readonly Dictionary<string, string> sessions = new();

    /// <param name="exePath">ExcelApp.exe, which the agent starts with --mcp.</param>
    /// <param name="find">Finds a CLI by name ("claude", "codex"); null means not installed. Tests pass their own.</param>
    public AgentRunner(string exePath, ConfigRoot config, Action<AgentEvent> emit, Func<string, string?>? find = null)
    { this.exePath = exePath; this.config = config; this.emit = emit; this.find = find ?? FindOnPath; }

    public void Register(Dictionary<string, Func<JsonElement, Task<object?>>> h)
    {
        h["agent.status"] = _ => Task.FromResult<object?>(Status());
        h["agent.start"] = a => Task.FromResult<object?>(Start(Arg.Str(a, "agent") ?? "claude", Arg.Str(a, "prompt") ?? "", Arg.Bool(a, "newChat")));
        h["agent.stop"] = _ => Task.FromResult<object?>(Stop());
    }

    public object Status()
    {
        lock (gate) return new { claude = find("claude") != null, codex = find("codex") != null, running = Running ? running : null };
    }
    bool Running => proc is { HasExited: false };

    // ---------------- finding and starting the CLI ----------------
    /// <summary>The CLI on PATH (Windows: .exe first, then the npm .cmd shim), or in the usual install folders.</summary>
    public static string? FindOnPath(string name)
    {
        var env = Environment.GetEnvironmentVariable("EXCELAPP_" + name.ToUpperInvariant());   // e.g. EXCELAPP_CLAUDE=D:\tools\claude.exe
        if (!string.IsNullOrWhiteSpace(env) && File.Exists(env)) return env;
        var dirs = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries).ToList();
        string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        dirs.Add(Path.Combine(home, ".local", "bin"));
        if (OperatingSystem.IsWindows()) dirs.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm"));
        string[] names = OperatingSystem.IsWindows() ? new[] { name + ".exe", name + ".cmd" } : new[] { name };
        foreach (var n in names) foreach (var d in dirs)
            {
                try { var p = Path.Combine(d.Trim('"'), n); if (File.Exists(p)) return p; } catch (ArgumentException) { }
            }
        return null;
    }

    /// <summary>The program and arguments for one turn. The request itself goes to the CLI's input, never on the command line.</summary>
    public static (string File, List<string> Args) Command(string agent, string cli, string mcpConfig, string workDir, string exe, string? resume)
    {
        var a = new List<string>();
        if (agent == "claude")
        {
            a.AddRange(new[] { "-p", "--output-format", "stream-json", "--verbose", "--mcp-config", mcpConfig, "--strict-mcp-config",
                               "--allowedTools", "mcp__excelapp" });
            if (resume != null) a.AddRange(new[] { "--resume", resume });
        }
        else
        {
            a.AddRange(new[] { "exec", "--json", "--skip-git-repo-check", "-C", workDir,
                               "-c", "mcp_servers.excelapp.command='" + exe + "'", "-c", "mcp_servers.excelapp.args=['--mcp']", "-c", "approval_policy='never'" });
            if (resume != null) a.AddRange(new[] { "resume", resume });
            a.Add("-");   // read the request from the input
        }
        return (cli, a);
    }

    public object Start(string agent, string prompt, bool newChat)
    {
        if (agent is not ("claude" or "codex")) throw new UserError("Choose Claude Code or Codex.");
        if (string.IsNullOrWhiteSpace(prompt)) throw new UserError("Write what the agent should do.");
        lock (gate)
        {
            if (Running) throw new UserError("The agent is still working. Wait, or press Stop.");
            string cli = find(agent) ?? throw new UserError(agent == "claude"
                ? "Claude Code is not installed. Install it (npm install -g @anthropic-ai/claude-code), sign in once with \"claude\", then try again."
                : "Codex is not installed. Install it (npm install -g @openai/codex), sign in once with \"codex\", then try again.");
            if (newChat) sessions.Remove(agent);
            sessions.TryGetValue(agent, out var resume);

            // Settings for the CLI live with ExcelApp's configuration (no cell data): which MCP server to start, and the instructions.
            config.WriteAllText("agent/mcp.json", JsonSerializer.Serialize(new { mcpServers = new { excelapp = new { command = exePath, args = new[] { "--mcp" } } } }));
            string workDir = Path.GetDirectoryName(config.PathOf("agent/mcp.json"))!;
            var (file, args) = Command(agent, cli, config.PathOf("agent/mcp.json"), workDir, exePath, resume);
            // The first message of a conversation carries the instructions; the CLI keeps them for the rest of it.
            string input = resume == null ? Instructions + "\n\n" + prompt : prompt;

            var psi = new ProcessStartInfo { WorkingDirectory = workDir, UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
            if (file.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase) || file.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
            {
                // npm's shim: cmd.exe runs it. Every argument is ours (paths and fixed words), quoted; the request goes to stdin.
                psi.FileName = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
                psi.Arguments = "/d /s /c \"" + Quote(file) + " " + string.Join(" ", args.Select(Quote)) + "\"";
            }
            else { psi.FileName = file; foreach (var x in args) psi.ArgumentList.Add(x); }

            var p = new Process { StartInfo = psi, EnableRaisingEvents = true };
            try { p.Start(); }
            catch (Exception e) when (e is System.ComponentModel.Win32Exception or IOException) { throw new UserError("The agent could not start (" + e.Message + ")."); }
            proc = p; running = agent;
            var stdin = new StreamWriter(p.StandardInput.BaseStream, new UTF8Encoding(false));
            stdin.Write(input); stdin.Close();
            var errTail = new StringBuilder();
            _ = Task.Run(async () => { try { string? l; while ((l = await p.StandardError.ReadLineAsync()) != null) lock (errTail) { errTail.AppendLine(l); if (errTail.Length > 4000) errTail.Remove(0, errTail.Length - 4000); } } catch (Exception) { } });
            _ = Task.Run(() => Pump(agent, p, errTail));
            emit(new AgentEvent("status", agent == "claude" ? "Claude Code is working..." : "Codex is working..."));
            return new { started = true, agent, resumed = resume != null };
        }
    }
    static string Quote(string s) => "\"" + s.Replace("\"", "") + "\"";

    async Task Pump(string agent, Process p, StringBuilder errTail)
    {
        bool done = false, failed = false;
        try
        {
            string? line;
            while ((line = await p.StandardOutput.ReadLineAsync()) != null)
            {
                string? session = null;
                foreach (var e in agent == "claude" ? ParseClaude(line, out session) : ParseCodex(line, out session))
                {
                    if (e.Kind == "done") done = true; if (e.Kind == "error") failed = true;
                    emit(e);
                }
                if (session != null) lock (gate) sessions[agent] = session;
            }
            await p.WaitForExitAsync();
        }
        catch (Exception e) { emit(new AgentEvent("error", "The agent stopped (" + e.GetType().Name + ").")); failed = true; }
        if (!done && !failed)
        {
            string tail; lock (errTail) tail = errTail.ToString().Trim();
            bool stopped; lock (gate) stopped = proc != p;
            emit(stopped ? new AgentEvent("done", "Stopped.")
                : new AgentEvent("error", p.ExitCode == 0 ? "The agent ended without an answer." : "The agent stopped with an error." + (tail.Length > 0 ? " " + LastLines(tail, 6) : "")));
        }
        lock (gate) if (proc == p) { proc = null; running = null; }
        p.Dispose();
    }
    static string LastLines(string s, int n) => string.Join("\n", s.Split('\n').TakeLast(n)).Trim();

    public object Stop()
    {
        Process? p; lock (gate) { p = proc; proc = null; running = null; }
        if (p != null) try { if (!p.HasExited) p.Kill(entireProcessTree: true); } catch (Exception e) when (e is InvalidOperationException or System.ComponentModel.Win32Exception) { }
        return new { stopped = p != null };
    }
    public void Dispose() => Stop();

    // ---------------- reading the CLIs' JSON events ----------------
    static string Short(string s, int max = 160) { s = s.Replace("\r", " ").Replace("\n", " ").Trim(); return s.Length > max ? s[..max] + "..." : s; }
    static string ToolName(string n) => n.StartsWith("mcp__excelapp__", StringComparison.Ordinal) ? n[15..] : n.StartsWith("mcp__", StringComparison.Ordinal) ? n[5..] : n;
    /// <summary>The most telling argument of a tool call: the SQL, the range, the command.</summary>
    static string ToolArgs(JsonElement input)
    {
        if (input.ValueKind == JsonValueKind.String) { try { using var d = JsonDocument.Parse(input.GetString()!); return ToolArgs(d.RootElement.Clone()); } catch (JsonException) { return Short(input.GetString()!); } }
        if (input.ValueKind != JsonValueKind.Object) return "";
        foreach (var k in new[] { "sql", "range", "start", "command", "table" }) if (input.TryGetProperty(k, out var v) && v.ValueKind == JsonValueKind.String) return Short(v.GetString()!);
        return "";
    }

    /// <summary>Claude Code: claude -p --output-format stream-json --verbose.</summary>
    public static IEnumerable<AgentEvent> ParseClaude(string line, out string? session)
    {
        session = null; var list = new List<AgentEvent>();
        JsonDocument doc; try { doc = JsonDocument.Parse(line); } catch (JsonException) { if (line.Trim().Length > 0) list.Add(new("status", Short(line))); return list; }
        using (doc)
        {
            var o = doc.RootElement; string type = o.TryGetProperty("type", out var t) ? t.GetString() ?? "" : "";
            if (o.TryGetProperty("session_id", out var sid) && sid.ValueKind == JsonValueKind.String && type is "system" or "result") session = sid.GetString();
            switch (type)
            {
                case "system" when o.TryGetProperty("subtype", out var st) && st.GetString() == "init":
                    if (o.TryGetProperty("mcp_servers", out var servers) && servers.ValueKind == JsonValueKind.Array)
                        foreach (var s in servers.EnumerateArray())
                            if (s.TryGetProperty("name", out var n) && n.GetString() == "excelapp" && s.TryGetProperty("status", out var ss) && ss.GetString() is string status && status != "connected")
                                list.Add(new("error", "Claude Code could not start ExcelApp's tools (" + status + ")."));
                    break;
                case "assistant":
                    if (o.TryGetProperty("message", out var m) && m.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.Array)
                        foreach (var c in content.EnumerateArray())
                        {
                            string ct = c.TryGetProperty("type", out var x) ? x.GetString() ?? "" : "";
                            if (ct == "text" && c.TryGetProperty("text", out var tx) && tx.GetString() is { Length: > 0 } text) list.Add(new("text", text));
                            else if (ct == "tool_use") list.Add(new("tool", ToolName(c.GetProperty("name").GetString() ?? "") + (c.TryGetProperty("input", out var inp) && ToolArgs(inp) is { Length: > 0 } ar ? ": " + ar : "")));
                        }
                    break;
                case "user":
                    if (o.TryGetProperty("message", out var um) && um.TryGetProperty("content", out var uc) && uc.ValueKind == JsonValueKind.Array)
                        foreach (var c in uc.EnumerateArray())
                            if (c.TryGetProperty("type", out var x) && x.GetString() == "tool_result" && c.TryGetProperty("is_error", out var ie) && ie.ValueKind == JsonValueKind.True)
                                list.Add(new("toolError", Short(c.TryGetProperty("content", out var rc) ? rc.ValueKind == JsonValueKind.String ? rc.GetString()! : string.Concat(rc.EnumerateArray().Select(i => i.TryGetProperty("text", out var it) ? it.GetString() : "")) : "The tool failed.")));
                    break;
                case "result":
                    bool err = o.TryGetProperty("is_error", out var e) && e.ValueKind == JsonValueKind.True;
                    string res = o.TryGetProperty("result", out var r) && r.ValueKind == JsonValueKind.String ? r.GetString()! : "";
                    if (err) list.Add(new("error", res.Length > 0 ? Short(res, 600) : "Claude Code stopped with an error."));
                    else list.Add(new("done", o.TryGetProperty("duration_ms", out var dm) && dm.TryGetInt64(out long ms) ? $"Done in {ms / 1000.0:0.#} s." : "Done."));
                    break;
            }
        }
        return list;
    }

    /// <summary>Codex: codex exec --json.</summary>
    public static IEnumerable<AgentEvent> ParseCodex(string line, out string? session)
    {
        session = null; var list = new List<AgentEvent>();
        JsonDocument doc; try { doc = JsonDocument.Parse(line); } catch (JsonException) { if (line.Trim().Length > 0) list.Add(new("status", Short(line))); return list; }
        using (doc)
        {
            var o = doc.RootElement; string type = o.TryGetProperty("type", out var t) ? t.GetString() ?? "" : "";
            switch (type)
            {
                case "thread.started": if (o.TryGetProperty("thread_id", out var id)) session = id.GetString(); break;
                case "error":
                    string msg = o.TryGetProperty("message", out var em) ? em.GetString() ?? "" : "";
                    list.Add(new(msg.StartsWith("Reconnecting", StringComparison.Ordinal) ? "status" : "error", Short(msg, 300))); break;
                case "turn.failed":
                    list.Add(new("error", o.TryGetProperty("error", out var fe) && fe.TryGetProperty("message", out var fm) ? Short(fm.GetString() ?? "", 400) : "Codex stopped with an error.")); break;
                case "turn.completed": list.Add(new("done", "Done.")); break;
                case "item.started" or "item.completed":
                    if (!o.TryGetProperty("item", out var item)) break;
                    string it = item.TryGetProperty("type", out var itt) ? itt.GetString() ?? "" : "";
                    if (it == "agent_message" && type == "item.completed" && item.TryGetProperty("text", out var tx) && tx.GetString() is { Length: > 0 } text) list.Add(new("text", text));
                    else if (it == "mcp_tool_call" && type == "item.started")
                        list.Add(new("tool", (item.TryGetProperty("tool", out var tn) ? tn.GetString() : "tool") + (item.TryGetProperty("arguments", out var ta) && ToolArgs(ta) is { Length: > 0 } ar ? ": " + ar : "")));
                    else if (it == "mcp_tool_call" && type == "item.completed" && item.TryGetProperty("status", out var ts) && ts.GetString() == "failed")
                        list.Add(new("toolError", item.TryGetProperty("error", out var te) && te.ValueKind == JsonValueKind.Object && te.TryGetProperty("message", out var tm) ? Short(tm.GetString() ?? "") : "The tool failed."));
                    else if (it == "command_execution" && type == "item.started" && item.TryGetProperty("command", out var cmd)) list.Add(new("tool", "shell: " + Short(cmd.GetString() ?? "")));
                    else if (it == "error" && item.TryGetProperty("message", out var ie)) list.Add(new("status", Short(ie.GetString() ?? "", 300)));
                    break;
            }
        }
        return list;
    }
}
