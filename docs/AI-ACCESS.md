# AI agents in ExcelApp

The owner decided on 2026-11-14 that AI may see all the data, cell values included (README, "The rule"), and that the
agent is a CLI agent (Claude Code or Codex) connected to ExcelApp, with no API key. An agent gets:
- **the data store, read-only, through SQL:** no export and no opening of Excel files. It runs on the Data mode's engine,
  so a question reads only the columns it uses.
- **the workbook open in the window, read and write:** cells by Excel address, plus every command of the window.

## The Agent panel (inside ExcelApp)

Press **✦ Agent** in the title bar. Choose **Claude Code** or **Codex**, write what you want in your own words, and the
agent does it in ExcelApp. You can see each step it takes (which tool, which range or SQL), stop it at any moment, and
start a new conversation. ExcelApp holds no API key. It runs the agent CLI you have installed and signed in to on
this computer:

- Claude Code: `npm install -g @anthropic-ai/claude-code`, then run `claude` once in a terminal to sign in.
- Codex: `npm install -g @openai/codex`, then run `codex` once to sign in.

ExcelApp finds them on PATH, in `%APPDATA%\npm` or in `~\.local\bin`. To use another location, set `EXCELAPP_CLAUDE` or
`EXCELAPP_CODEX` to the full path. The agent gets only ExcelApp's tools (Claude Code runs with
`--strict-mcp-config --allowedTools mcp__excelapp`). Its changes appear in the grid at once, they are one undo step
each, and they stay unsaved until you save or ask the agent to save.

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

The server offers these tools:

| Tool | What it does |
|---|---|
| `list_tables` | Every report type and table in the data store, with row counts and columns |
| `describe_table` | One table's columns and types (number, text, date, bool), plus sample rows |
| `query` | The answer to one SQL `SELECT`: columns, rows (1,000 by default, up to 100,000) and how many rows matched |
| `app_status` | The workbook open in the ExcelApp window: name, unsaved changes, sheets and their used ranges |
| `read_range` | Cells of the open workbook as Excel shows them, with formulas (`Sheet1!A1:D20`, at most 20,000 cells) |
| `write_range` | Writes rows of values from a cell: numbers, text, true/false, null to clear, `=...` formulas. One undo step |
| `run_command` | Any command of the window: styles, filters, sorting, fill, find and replace, undo, save, the data engine's `data.*` commands |

The store tools work without the window. The workbook tools reach the open window through a named pipe that only your
Windows account can open, and they say "ExcelApp is not open" when no window is running.

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
