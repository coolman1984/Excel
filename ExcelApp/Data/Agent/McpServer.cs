// A Model Context Protocol server over stdio (JSON-RPC 2.0, one message per line) so an AI agent such as Claude
// can list, describe and query the store directly. Read-only. No packages: plain System.Text.Json.
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ExcelApp.Data.Agent;

public sealed class McpServer
{
    readonly AgentData data;
    public McpServer(AgentData data) => this.data = data;

    static readonly string[] Versions = { "2025-06-18", "2025-03-26", "2024-11-05" };
    const string Instructions =
        "ExcelApp's data store: report types (every loaded period of a report, latest version) and tables. " +
        "Start with list_tables, then describe_table for column names and sample rows, then query with SQL. " +
        "SQL: SELECT [DISTINCT] columns, expressions (+ - * /, ABS, ROUND, COALESCE, MONTH, PCT) and SUM/COUNT/COUNT(DISTINCT)/AVG/MIN/MAX " +
        "FROM table [[LEFT] JOIN table ON a.col = b.col] [WHERE ... with = <> < > <= >=, AND/OR/NOT, IN (...), BETWEEN, LIKE 'text%', IS [NOT] NULL] " +
        "[GROUP BY cols] [HAVING ...] [ORDER BY col|alias|position [DESC]] [LIMIT n]. Quote names with spaces as \"Unit price\" or [Unit price]; " +
        "text values in single quotes; dates as 'yyyy-mm-dd'. Text comparisons ignore case. Answers are capped (max_rows, default 1000); rowCount says how many rows matched.";

    static JsonObject Tool(string name, string description, JsonObject properties, params string[] required) => new()
    {
        ["name"] = name, ["description"] = description,
        ["inputSchema"] = new JsonObject { ["type"] = "object", ["properties"] = properties, ["required"] = new JsonArray(required.Select(r => (JsonNode)r!).ToArray()) },
        ["annotations"] = new JsonObject { ["readOnlyHint"] = true, ["openWorldHint"] = false },
    };

    static JsonArray ToolList() => new()
    {
        Tool("list_tables", "Lists every table and report type in ExcelApp's store with its row count and columns.", new JsonObject()),
        Tool("describe_table", "Shows a table's columns with their types (number, text, date, bool) and a few sample rows.",
            new JsonObject { ["table"] = new JsonObject { ["type"] = "string", ["description"] = "Table or report type name from list_tables." },
                             ["sample_rows"] = new JsonObject { ["type"] = "integer", ["description"] = "Sample rows to include (0 to 50, default 5)." } }, "table"),
        Tool("query", "Runs one read-only SQL SELECT on the store and returns columns, rows and the matched row count. " + Instructions,
            new JsonObject { ["sql"] = new JsonObject { ["type"] = "string", ["description"] = "One SELECT statement." },
                             ["max_rows"] = new JsonObject { ["type"] = "integer", ["description"] = $"Most rows to return (default {AgentData.DefaultRows}, at most {AgentData.MaxRows})." } }, "sql"),
    };

    /// <summary>Answers one message; null for a notification (no id).</summary>
    public string? Handle(string line)
    {
        JsonNode? id = null;
        try
        {
            JsonNode? msg;
            try { msg = JsonNode.Parse(line); } catch (JsonException) { return Error(null, -32700, "The message is not valid JSON."); }
            if (msg is not JsonObject o || o["method"]?.GetValue<string>() is not string method) return Error(null, -32600, "Not a JSON-RPC request.");
            bool notification = !o.ContainsKey("id");
            id = o["id"]?.DeepClone();
            var p = o["params"] as JsonObject;
            if (notification) return null;
            switch (method)
            {
                case "initialize":
                    string asked = p?["protocolVersion"]?.GetValue<string>() ?? Versions[0];
                    return Result(id, new JsonObject
                    {
                        ["protocolVersion"] = Versions.Contains(asked) ? asked : Versions[0],
                        ["capabilities"] = new JsonObject { ["tools"] = new JsonObject { ["listChanged"] = false } },
                        ["serverInfo"] = new JsonObject { ["name"] = "excelapp", ["title"] = "ExcelApp data store", ["version"] = "1.0" },
                        ["instructions"] = Instructions,
                    });
                case "ping": return Result(id, new JsonObject());
                case "tools/list": return Result(id, new JsonObject { ["tools"] = ToolList() });
                case "tools/call": return Result(id, Call(p?["name"]?.GetValue<string>() ?? "", p?["arguments"] as JsonObject ?? new JsonObject()));
                default: return Error(id, -32601, "Unknown method: " + method);
            }
        }
        catch (Exception e) { return Error(id, -32603, "Internal error (" + e.GetType().Name + ")."); }
    }

    JsonObject Call(string name, JsonObject args)
    {
        try
        {
            object answer = name switch
            {
                "list_tables" => data.Tables(),
                "describe_table" => data.Describe(Str(args, "table"), Int(args, "sample_rows") ?? 5),
                "query" => data.Query(Str(args, "sql"), Int(args, "max_rows") ?? AgentData.DefaultRows),
                _ => throw new UserError("Unknown tool: " + name + ". Tools: list_tables, describe_table, query."),
            };
            return Content(JsonSerializer.Serialize(answer, AgentData.JsonOptions), false);
        }
        catch (UserError e) { return Content(e.Message, true); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or InvalidDataException) { return Content("The store could not be read (" + e.GetType().Name + "). Try again in a moment.", true); }
    }
    static string Str(JsonObject a, string key) => a[key] is JsonValue v && v.TryGetValue<string>(out var s) ? s : throw new UserError($"The argument \"{key}\" is missing.");
    static int? Int(JsonObject a, string key) => a[key] is JsonValue v && v.TryGetValue<int>(out var n) ? n : a[key] is JsonValue d && d.TryGetValue<double>(out var x) ? (int)x : null;
    static JsonObject Content(string text, bool error) => new() { ["content"] = new JsonArray(new JsonObject { ["type"] = "text", ["text"] = text }), ["isError"] = error };

    static string Result(JsonNode? id, JsonNode result) => new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result }.ToJsonString();
    static string Error(JsonNode? id, int code, string message) => new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = new JsonObject { ["code"] = code, ["message"] = message } }.ToJsonString();

    /// <summary>Serves until stdin closes. Only protocol messages go to stdout.</summary>
    public void Run(TextReader input, TextWriter output)
    {
        string? line;
        while ((line = input.ReadLine()) != null)
        {
            if (line.Length == 0) continue;
            var reply = Handle(line);
            if (reply != null) { output.WriteLine(reply); output.Flush(); }
        }
    }
}
