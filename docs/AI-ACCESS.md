# AI access to the data store

The owner decided on 2026-11-14 that AI may read all the data, cell values included (README, "The rule").
ExcelApp gives an AI agent **direct, read-only** access to its store, with no export and no opening of Excel files.
It speaks SQL, and it runs on the same engine as the Data mode, so a question reads only the columns it uses.

## Connect an AI app (MCP)

`ExcelApp.exe --mcp` is a [Model Context Protocol](https://modelcontextprotocol.io) server on stdio. Build the app once
(`ExcelApp\build.ps1`, or start `ExcelApp\Start_ExcelApp.bat`), then point the AI app at `ExcelApp\app\ExcelApp.exe`.
Replace `C:\path\to\Excel` below with the folder that holds this repository.

**Claude Desktop:** add this to `%APPDATA%\Claude\claude_desktop_config.json`, then restart Claude Desktop:
```json
{
  "mcpServers": {
    "excelapp": { "command": "C:\\path\\to\\Excel\\ExcelApp\\app\\ExcelApp.exe", "args": ["--mcp"] }
  }
}
```

**Claude Code:**
```
claude mcp add excelapp -- "C:\path\to\Excel\ExcelApp\app\ExcelApp.exe" --mcp
```

The server offers three tools, all read-only:

| Tool | What it returns |
|---|---|
| `list_tables` | Every report type and table, with row counts and columns |
| `describe_table` | One table's columns and types (number, text, date, bool), plus sample rows |
| `query` | The answer to one SQL `SELECT`: columns, rows (1,000 by default, up to 100,000) and how many rows matched |

The app does not have to be open. The server reads the store files directly, and the app can keep loading new files
while an agent reads.

## From a terminal or a script

```
ExcelApp.exe --tables
ExcelApp.exe --query "SELECT Region, SUM(Sales) AS Total FROM sales GROUP BY Region ORDER BY Total DESC" --max 100
```
Each prints one line of JSON. On a problem it prints `ERROR <sentence>` and exits with code 1.

## The SQL

```sql
SELECT [DISTINCT] col, col AS name, expression AS name, SUM(col), COUNT(*), COUNT(DISTINCT col), AVG(col), MIN(col), MAX(col)
FROM table [alias]
[[LEFT] JOIN other ON table.key = other.key [AND ...]]
[WHERE  col = 'text' AND (col > 10 OR col IN ('a','b')) AND col BETWEEN 1 AND 5
        AND col LIKE 'abc%' AND col IS NOT NULL AND NOT col = 3]
[GROUP BY col, ...] [HAVING SUM(col) > 100]
[ORDER BY col | alias | position [DESC], ...] [LIMIT n]
```
- **Tables:** a report type name means every loaded period of that report, latest version. A table name means that table.
  `list_tables` shows both.
- **Names** with spaces or Arabic letters go in `"double quotes"` or `[brackets]`, and text values go in `'single quotes'`.
  Names and text comparisons ignore upper and lower case.
- **Calculations:** `+ - * /` and the functions `ABS`, `ROUND(x, digits)`, `COALESCE(x, y)`, `MONTH(date)` and `PCT(part, whole)`.
- **Dates** compare with `'yyyy-mm-dd'` text and come back in the same form.
- **Joins:** columns of a joined table are named `table.column`. Write `targets.Target` when both tables have the column.
- **LIKE** understands `text%`, `%text` and `%text%` only.
- **Not allowed:** anything other than `SELECT`, subqueries, `UNION`, window functions, and arithmetic on aggregates
  (write `SUM(a)` and `SUM(b)` as two columns instead of `SUM(a)/SUM(b)`).

## Speed (made-up data, 3 months x 1,000,000 rows x 300 columns, over MCP)

Measured on the Linux build machine; re-measure on the laptop.

| Question | Time |
|---|---|
| `list_tables` | 0.08 s |
| `describe_table` (300 columns, sample rows) | 0.2 s |
| Sum and count by one of 2,000 regions, top 3 | 1.1 s |
| Count, average and max with a number filter and a `LIKE` | 0.75 s |
| A calculated column sorted over 3,000,000 rows, top 2 | 1.5 s |

The server process peaked at 860 MB across all of these. `SELECT * ... LIMIT n` filters and sorts on the columns it
needs, then reads the 300 columns only for the n rows it returns.
