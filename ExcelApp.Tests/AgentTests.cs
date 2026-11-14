// AI agent access: the SQL-like language over the store, the read-only agent API and the MCP server. Made-up data only.
using System.Text.Json;
using ExcelApp.Data;
using ExcelApp.Data.Agent;

public static class AgentTests
{
    sealed class Store : ITableProvider
    {
        public readonly Dictionary<string, Table> Tables = new(StringComparer.OrdinalIgnoreCase);
        public Table Get(TableRef r) => Tables.TryGetValue(r.Id, out var t) ? t : throw new ExcelApp.UserError("The table was not found.");
        public IReadOnlyList<TableInfo> List() => Tables.Select(kv => new TableInfo(kv.Key == "sales" ? "report" : "table", kv.Key, kv.Key, "", kv.Value.Rows, kv.Value.Columns.Select(c => new ColumnInfo(c.Name, c.Type)).ToArray(), new DateTime(2026, 11, 1))).ToArray();
        public SourceInfo? SourceOf(int id) => null;
        public int SaveAsVersion(string reportTypeId, string period, Table data, string note) => throw new NotSupportedException();
    }

    public static void Run(Action<string, bool, string> check)
    {
        var store = new Store();
        store.Tables["sales"] = new Table("sales", new Column[] {
            TextColumn.From("Region", new[] { "North", "South", "North", "East", null }),
            new NumberColumn("Sales", new[] { 10d, 20, 30, 40, 50 }, null),
            new NumberColumn("Cost", new[] { 4d, 5, 6, 7, 0 }, new[] { false, false, false, false, true }),
            TextColumn.From("Rep", new[] { "O'Brien", "Ann", "Bo", "Ann", "Cy" }),
            new NumberColumn("Day", new[] { 46000d, 46001, 46002, 46003, 46004 }, null, ColType.Date),
            new NumberColumn("_v", new[] { 1d, 1, 1, 1, 1 }, null) });
        store.Tables["targets"] = new Table("targets", new Column[] { TextColumn.From("Region", new[] { "North", "South" }), new NumberColumn("Target", new[] { 35d, 15 }, null) });
        var api = new AgentData(store);

        AgentResult Q(string sql, int max = 1000) => api.Query(sql, max);
        string Err(string sql) { try { Q(sql); return "no error"; } catch (ExcelApp.UserError e) { return e.Message; } catch (Exception e) { return "RAW " + e.GetType().Name; } }
        object? Cell(AgentResult r, int row, string col) => r.Rows[row][r.Columns.ToList().FindIndex(c => c.Name == col)];
        string Names(AgentResult r) => string.Join(",", r.Columns.Select(c => c.Name));

        var g = Q("SELECT Region, SUM(Sales) AS Total, COUNT(*) AS N FROM sales WHERE Sales > 5 GROUP BY Region ORDER BY Total DESC");
        check("sql: group, sum, count, order", Names(g) == "Region,Total,N" && (string?)Cell(g, 0, "Region") == null && (double)Cell(g, 0, "Total")! == 50 && (string?)Cell(g, 1, "Region") == "North" && (double)Cell(g, 1, "Total")! == 40 && (double)Cell(g, 1, "N")! == 2, Names(g) + " " + JsonSerializer.Serialize(g.Rows));
        var agg = Q("select count(distinct Region) as Regions, avg(Sales) as A, min(Sales) as Lo, max(Sales) as Hi, count(Cost) as WithCost from sales");
        check("sql: whole-table aggregates", (double)Cell(agg, 0, "Regions")! == 3 && (double)Cell(agg, 0, "A")! == 30 && (double)Cell(agg, 0, "Lo")! == 10 && (double)Cell(agg, 0, "Hi")! == 50 && (double)Cell(agg, 0, "WithCost")! == 4, JsonSerializer.Serialize(agg.Rows));
        check("sql: default aggregate names", Names(Q("SELECT Region, sum(Sales) FROM sales GROUP BY Region")) == "Region,SUM(Sales)", Names(Q("SELECT Region, sum(Sales) FROM sales GROUP BY Region")));
        check("sql: AND, OR, NOT and brackets", Q("SELECT Sales FROM sales WHERE (Region = 'North' OR Region = 'East') AND NOT Sales = 30").RowCount == 2, "");
        check("sql: IN and BETWEEN", Q("SELECT * FROM sales WHERE Region IN ('North', 'South') AND Sales BETWEEN 15 AND 30").RowCount == 2, "");
        check("sql: LIKE begins, ends, contains", Q("SELECT * FROM sales WHERE Region LIKE 'no%'").RowCount == 2 && Q("SELECT * FROM sales WHERE Region LIKE '%th'").RowCount == 3 && Q("SELECT * FROM sales WHERE Region LIKE '%as%'").RowCount == 1, "");
        check("sql: IS NULL and IS NOT NULL", Q("SELECT * FROM sales WHERE Region IS NULL").RowCount == 1 && Q("SELECT * FROM sales WHERE Cost IS NOT NULL").RowCount == 4, "");
        check("sql: quoted text with an apostrophe", Q("SELECT * FROM sales WHERE Rep = 'O''Brien'").RowCount == 1, "");
        check("sql: dates compare with ISO text", Q("SELECT * FROM sales WHERE Day >= '2025-12-12'").RowCount == 2 && Q("SELECT * FROM sales WHERE Day < DATE '2025-12-10'").RowCount == 1, Q("SELECT * FROM sales WHERE Day >= '2025-12-12'").RowCount.ToString());
        var date = Q("SELECT Day FROM sales ORDER BY Day LIMIT 1");
        check("sql: dates come back as ISO text", (string?)Cell(date, 0, "Day") == "2025-12-09", Cell(date, 0, "Day")?.ToString() ?? "null");
        var calc = Q("SELECT Rep, Sales * 2 AS Twice, Sales - Cost AS Margin FROM sales ORDER BY Twice DESC LIMIT 2");
        check("sql: arithmetic, order by alias, limit", calc.RowCount == 2 && (double)Cell(calc, 0, "Twice")! == 100 && Cell(calc, 0, "Margin") == null && (double)Cell(calc, 1, "Margin")! == 33, JsonSerializer.Serialize(calc.Rows));
        var margin = Q("SELECT SUM(Sales - Cost) AS Margin FROM sales WHERE Cost IS NOT NULL");
        check("sql: aggregate over an expression", (double)Cell(margin, 0, "Margin")! == 78, JsonSerializer.Serialize(margin.Rows));
        check("sql: DISTINCT", Q("SELECT DISTINCT Rep FROM sales").RowCount == 4, "");
        check("sql: HAVING on an aggregate", Q("SELECT Region, SUM(Sales) AS T FROM sales GROUP BY Region HAVING SUM(Sales) > 35").RowCount == 3, Q("SELECT Region, SUM(Sales) AS T FROM sales GROUP BY Region HAVING SUM(Sales) > 35").RowCount.ToString());
        check("sql: ORDER BY position and several keys", (double)Cell(Q("SELECT Rep, Sales FROM sales ORDER BY 1, 2 DESC"), 0, "Sales")! == 40, "");
        check("sql: SELECT * hides lineage columns", Names(Q("SELECT * FROM sales")) == "Region,Sales,Cost,Rep,Day", Names(Q("SELECT * FROM sales")));
        var join = Q("SELECT sales.Region, SUM(Sales) AS S, MAX(targets.Target) AS T FROM sales JOIN targets ON sales.Region = targets.Region GROUP BY sales.Region ORDER BY S DESC");
        check("sql: JOIN with table-qualified names", join.RowCount == 2 && (string?)Cell(join, 0, "Region") == "North" && (double)Cell(join, 0, "T")! == 35, Names(join) + " " + JsonSerializer.Serialize(join.Rows));
        check("sql: LEFT JOIN keeps unmatched rows", Q("SELECT * FROM sales LEFT JOIN targets ON Region = Region").RowCount == 5, "");
        store.Tables["noise"] = new Table("noise", new Column[] { new NumberColumn("X", new[] { 0.1, 0.2 }, null) });
        check("agent: sums show 15 significant digits like Excel", (double)Cell(Q("SELECT SUM(X) AS S FROM noise"), 0, "S")! == 0.3, Cell(Q("SELECT SUM(X) AS S FROM noise"), 0, "S")?.ToString() ?? "null");
        var cut = Q("SELECT * FROM sales", 2);
        check("sql: answers are capped and say so", cut.Rows.Count == 2 && cut.RowCount == 5 && cut.Truncated, "");
        check("sql: names are case-insensitive, trailing ; allowed", Q("select region from SALES where sales > 25;").RowCount == 3, "");

        check("sql error: unknown table names the tables", Err("SELECT * FROM nothing").Contains("sales") && Err("SELECT * FROM nothing").Contains("targets"), Err("SELECT * FROM nothing"));
        check("sql error: unknown column is a plain sentence", !Err("SELECT Nope FROM sales").StartsWith("RAW") && Err("SELECT Nope FROM sales").Contains("Nope"), Err("SELECT Nope FROM sales"));
        check("sql error: syntax names the place", Err("SELECT FROM sales").Contains("FROM"), Err("SELECT FROM sales"));
        check("sql error: only SELECT is allowed", Err("DELETE FROM sales").Contains("SELECT"), Err("DELETE FROM sales"));
        check("sql error: an unsupported LIKE pattern is refused plainly", !Err("SELECT * FROM sales WHERE Rep LIKE 'a%b'").StartsWith("RAW") && Err("SELECT * FROM sales WHERE Rep LIKE 'a%b'") != "no error", Err("SELECT * FROM sales WHERE Rep LIKE 'a%b'"));
        check("sql error: unclosed text", !Err("SELECT * FROM sales WHERE Rep = 'abc").StartsWith("RAW") && Err("SELECT * FROM sales WHERE Rep = 'abc") != "no error", Err("SELECT * FROM sales WHERE Rep = 'abc"));

        var tables = JsonSerializer.Serialize(api.Tables(), AgentData.JsonOptions);
        check("agent: tables list names, rows and columns", tables.Contains("\"sales\"") && tables.Contains("\"rows\":5") && tables.Contains("\"Region\"") && !tables.Contains("\"_v\""), tables);
        store.Tables["arabic"] = new Table("arabic", new Column[] { TextColumn.From("المنطقة", new[] { "القاهرة" }) });
        var arabic = JsonSerializer.Serialize(api.Query("SELECT [المنطقة] FROM arabic"), AgentData.JsonOptions);
        check("agent: Arabic names and values come back readable", arabic.Contains("القاهرة") && arabic.Contains("المنطقة"), arabic);
        var described = JsonSerializer.Serialize(api.Describe("sales", 2), AgentData.JsonOptions);
        check("agent: describe shows types and sample rows", described.Contains("\"date\"") && described.Contains("O'Brien") && !described.Contains("\"Cy\""), described);

        // MCP over stdio: JSON-RPC 2.0, one message per line
        var mcp = new McpServer(api);
        JsonElement Call(string json) => JsonDocument.Parse(mcp.Handle(json) ?? "null").RootElement;
        var init = Call("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{\"protocolVersion\":\"2025-06-18\",\"capabilities\":{},\"clientInfo\":{\"name\":\"t\",\"version\":\"1\"}}}");
        check("mcp: initialize", init.GetProperty("id").GetInt32() == 1 && init.GetProperty("result").GetProperty("serverInfo").GetProperty("name").GetString() == "excelapp" && init.GetProperty("result").GetProperty("capabilities").TryGetProperty("tools", out _), init.ToString());
        check("mcp: a notification gets no answer", mcp.Handle("{\"jsonrpc\":\"2.0\",\"method\":\"notifications/initialized\"}") == null, "");
        var list = Call("{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"tools/list\"}");
        var toolNames = list.GetProperty("result").GetProperty("tools").EnumerateArray().Select(t => t.GetProperty("name").GetString()).ToArray();
        check("mcp: three read-only tools", toolNames.SequenceEqual(new[] { "list_tables", "describe_table", "query" }), string.Join(",", toolNames));
        var q = Call("{\"jsonrpc\":\"2.0\",\"id\":\"q1\",\"method\":\"tools/call\",\"params\":{\"name\":\"query\",\"arguments\":{\"sql\":\"SELECT Region, SUM(Sales) AS T FROM sales GROUP BY Region ORDER BY T DESC\"}}}");
        var text = q.GetProperty("result").GetProperty("content")[0].GetProperty("text").GetString()!;
        check("mcp: query answers as JSON text", q.GetProperty("id").GetString() == "q1" && !q.GetProperty("result").GetProperty("isError").GetBoolean() && text.Contains("\"North\"") && text.Contains("\"rowCount\":4"), text);
        var bad = Call("{\"jsonrpc\":\"2.0\",\"id\":3,\"method\":\"tools/call\",\"params\":{\"name\":\"query\",\"arguments\":{\"sql\":\"SELECT * FROM nothing\"}}}");
        check("mcp: a bad query is a tool error with the sentence", bad.GetProperty("result").GetProperty("isError").GetBoolean() && bad.GetProperty("result").GetProperty("content")[0].GetProperty("text").GetString()!.Contains("sales"), bad.ToString());
        var unknown = Call("{\"jsonrpc\":\"2.0\",\"id\":4,\"method\":\"nope\"}");
        check("mcp: unknown method error", unknown.GetProperty("error").GetProperty("code").GetInt32() == -32601, unknown.ToString());
        var broken = Call("{not json");
        check("mcp: unreadable message is a parse error", broken.GetProperty("error").GetProperty("code").GetInt32() == -32700, broken.ToString());
        check("mcp: ping", Call("{\"jsonrpc\":\"2.0\",\"id\":5,\"method\":\"ping\"}").TryGetProperty("result", out _), "");
    }
}
