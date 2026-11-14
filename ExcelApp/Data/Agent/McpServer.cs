// A Model Context Protocol server over stdio (JSON-RPC 2.0, one message per line) so an AI agent such as Claude
// can list, describe and query the store directly. Read-only. No packages: plain System.Text.Json.
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ExcelApp.Data.Agent;

public sealed class McpServer
{
    readonly AgentData? data; readonly ExcelApp.Agent.IAppLink? app; readonly string storeProblem;
    /// <summary><paramref name="data"/> answers the store tools (null when the store cannot open here, with the reason);
    /// <paramref name="app"/> reaches the running ExcelApp window for the workbook tools (null leaves them out).</summary>
    public McpServer(AgentData? data, ExcelApp.Agent.IAppLink? app = null, string? storeProblem = null)
    { this.data = data; this.app = app; this.storeProblem = storeProblem ?? "The data store cannot be opened on this computer."; }

    static readonly string[] Versions = { "2025-06-18", "2025-03-26", "2024-11-05" };
    const string AppInstructions =
        "The open workbook in the ExcelApp window: app_status shows the file, its sheets and their used ranges; read_range and write_range " +
        "use Excel addresses (Sheet1!A1:D20); write_range takes rows of values where text starting with = is a formula, and is one undo step. " +
        "run_command runs any of the window's own commands (styles, filters, sorting, fill, find and replace, undo, save, and the data engine's data.* commands). ";
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

    JsonArray ToolList()
    {
        var tools = StoreTools();
        if (app != null) foreach (var t in AppTools()) tools.Add(t!.DeepClone());
        return tools;
    }
    static JsonObject Str(string description) => new() { ["type"] = "string", ["description"] = description };
    static JsonArray AppTools() => new()
    {
        Tool("app_status", "The workbook open in the ExcelApp window: file name, unsaved changes, the current sheet, and every sheet's used range, rows and columns. " + AppInstructions, new JsonObject()),
        Tool("read_range", "Reads cells of the open workbook as Excel shows them (formatted text), with the formula of any formula cell. At most 20,000 cells per call.",
            new JsonObject { ["range"] = Str("Excel range, such as A1:D20, Sheet1!B2:B50, C:C or 3:5."), ["sheet"] = Str("Sheet name, when the range does not name one. Default: the current sheet.") }, "range"),
        Write("write_range", "Writes values into the open workbook starting at one cell, row by row, as typed into Excel: numbers, text, true/false, null to clear, and text starting with = as a formula. One undo step; the file is changed only when saved.",
            new JsonObject { ["start"] = Str("Top-left cell, such as B2 or Sheet1!B2."), ["sheet"] = Str("Sheet name, when start does not name one."),
                             ["values"] = new JsonObject { ["type"] = "array", ["description"] = "Rows of values, such as [[\"Total\", \"=SUM(B2:B9)\"]].", ["items"] = new JsonObject { ["type"] = "array" } } }, "start", "values"),
        Write("run_command", "Runs one of the ExcelApp window's own commands. Coordinates here are 0-based inside the sheet's data (row 0 = the sheet's first data row, col 0 = its first column, as app_status shows), and a filtered or sorted sheet counts visible rows. " +
            "Commands: sheet {sheet}; setStyle {sheet, ranges:[[r1,c1,r2,c2]], style:{bold, italic, underline, fill:\"#RRGGBB\", fontColor:\"#RRGGBB\", numberFormat, align, border:\"all|outside|bottom|none\"}}; " +
            "clearCells {sheet, ranges, what:\"contents|formats|all\"}; fill {sheet, from:[r1,c1,r2,c2], to:[r1,c1,r2,c2], copy}; " +
            "setFilter {sheet, col, condition:{kind:\"values\", include:[...], blanks} or {kind:\"number\", op:\"gt|ge|lt|le|eq|ne|between\", a, b} or {kind:\"text\", op:\"contains|begins|ends|eq|ne\", value}}; clearFilters {sheet}; " +
            "setSort {sheet, keys:[{col, desc}]}; find {query, matchCase, sheet}; replace {sheet, query, replacement, matchCase, all}; setColWidth {sheet, col, px}; setFreeze {sheet, rows, cols}; " +
            "data.dashboard.list; data.dashboard.save {dashboard:{id, name, tiles:[{id, title, kind:\"kpi|chart|table\", chartType, sql:\"SELECT ...\", format:\"number|currency|percent\", x, y, w, h}]}} (a 12-column grid; x 0-11, w 1-12, rows of 84 px); data.dashboard.values {id}; " +
            "pivot.fields {sheet, address}; pivot.run {sheet, address:\"A1:F500\", rows:[field], column:field, values:[{field, fn:\"sum|count|avg|min|max\"}]}; pivot.place {same} (writes it to the right of the data); " +
            "chart.add {sheet, address:\"A1:C13\" (header row and numbers), type:\"column|bar|line|area|pie|scatter\", title}; chart.update {id, type, title, x, y, w, h}; chart.delete {id}; chart.list {sheet}; " +
            "insertRows {sheet, row, count}; deleteRows {sheet, row, count}; insertCols {sheet, col, count}; deleteCols {sheet, col, count} (references move as in Excel; the sheet must have no filter or sort); " +
            "undo {sheet}; redo {sheet}; save {saveAs:false}; and the data engine's commands, such as data.views.list, data.recipes.run {id}, data.dashboards.list.",
            new JsonObject { ["command"] = Str("Command name, such as setSort."), ["args"] = new JsonObject { ["type"] = "object", ["description"] = "The command's arguments." } }, "command"),
    };
    static JsonObject Write(string name, string description, JsonObject properties, params string[] required)
    { var t = Tool(name, description, properties, required); t["annotations"] = new JsonObject { ["readOnlyHint"] = false, ["destructiveHint"] = false, ["openWorldHint"] = false }; return t; }

    static JsonArray StoreTools() => new()
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
            AgentData Store() => data ?? throw new UserError(storeProblem);
            ExcelApp.Agent.IAppLink App() => app ?? throw new UserError("Unknown tool: " + name + ".");
            object answer = name switch
            {
                "list_tables" => Store().Tables(),
                "describe_table" => Store().Describe(Text(args, "table"), Int(args, "sample_rows") ?? 5),
                "query" => Store().Query(Text(args, "sql"), Int(args, "max_rows") ?? AgentData.DefaultRows),
                "app_status" => Status(App().Call("workbook", Empty())),
                "read_range" => App().Call("agent.read", Where(args, "range", "range")),
                "write_range" => App().Call("agent.write", Where(args, "start", "start", ("values", args["values"] ?? throw new UserError("The argument \"values\" is missing.")))),
                "run_command" => App().Call(Text(args, "command"), JsonSerializer.SerializeToElement(args["args"] as JsonObject ?? new JsonObject())),
                _ => throw new UserError("Unknown tool: " + name + "."),
            };
            return Content(JsonSerializer.Serialize(answer, AgentData.JsonOptions), false);
        }
        catch (UserError e) { return Content(e.Message, true); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or InvalidDataException) { return Content("The store could not be read (" + e.GetType().Name + "). Try again in a moment.", true); }
    }
    static JsonElement Empty() => JsonSerializer.SerializeToElement(new { });
    /// <summary>Splits "Sheet1!B2" into the sheet and the address; a separate "sheet" argument also works.</summary>
    static JsonElement Where(JsonObject args, string key, string name, params (string Key, JsonNode Value)[] more)
    {
        string text = Text(args, key); string? sheet = args["sheet"] is JsonValue sv && sv.TryGetValue<string>(out var sn) ? sn : null;
        int bang = text.LastIndexOf('!');
        if (bang > 0) { sheet = text[..bang].Trim().Trim('\'').Replace("''", "'"); text = text[(bang + 1)..]; }
        var o = new JsonObject { [name] = text };
        if (sheet != null) o["sheet"] = sheet;
        foreach (var (k, v) in more) o[k] = v.DeepClone();
        return JsonSerializer.SerializeToElement(o);
    }
    /// <summary>The workbook summary an agent needs: no per-column widths.</summary>
    static object Status(JsonElement w)
    {
        if (w.ValueKind != JsonValueKind.Object) return new { open = false };
        var sheets = w.GetProperty("sheets").EnumerateArray().Select(s =>
        {
            int rows = s.GetProperty("rows").GetInt32(), cols = s.GetProperty("cols").GetInt32(), fr = s.GetProperty("firstRow").GetInt32(), fc = s.GetProperty("firstCol").GetInt32();
            return new { index = s.GetProperty("index").GetInt32(), name = s.GetProperty("name").GetString(), rows, cols, firstRow = fr, firstCol = fc,
                         used = rows == 0 || cols == 0 ? "" : ExcelApp.Agent.A1.Format(fr, fc, fr + rows - 1, fc + cols - 1), hidden = s.GetProperty("hidden").GetBoolean() };
        }).ToArray();
        return new { open = true, name = w.GetProperty("name").GetString(), dirty = w.GetProperty("dirty").GetBoolean(), currentSheet = w.GetProperty("currentSheet").GetInt32(), sheets };
    }
    static string Text(JsonObject a, string key) => a[key] is JsonValue v && v.TryGetValue<string>(out var s) ? s : throw new UserError($"The argument \"{key}\" is missing.");
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
