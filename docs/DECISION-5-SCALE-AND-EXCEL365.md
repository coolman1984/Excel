# Decision 5: 1,000,000 rows x 300 columns, and how far ExcelApp is from Excel 365

Date: 2026-11-14. Made-up data only (`syn-*`), measured with the editor's and the store's own code on a Linux build machine
(4 CPUs, 15 GB RAM, warm disk cache). The laptop (8 CPUs, 15.7 GB) has a colder disk: expect the first read of a part
of the file to be slower there. Re-measure on the laptop before quoting these numbers to anyone.

## The owner's target
One program that is Excel 365, a database and a dashboard at once. It handles more than 1,000,000 rows and more than 300
columns smoothly, with little RAM. AI reads its data directly, like SQL. Folders, reports and dashboards all live inside it.

## Measured: the editor (fast copy, 300,000,000 cells, 2.8 GB on disk)
| Action | Time | Memory of the whole program |
|---|---|---|
| Open (read the directory) | 0.1 s | 69 MB |
| First screen, 40 rows x 30 columns | 0.01 s | |
| Jump to any screen (median / 95%) | 0.3 ms / 0.4 ms | 95 MB |
| Scroll | 5 million cells per second | |
| Filter one number column | 0.31 s | 161 MB |
| Filter two columns (number and text) | 0.67 s | |
| Sort 1,000,000 rows by a number | 0.39 s | 256 MB |
| Sort by text, then number | 2.2 s | 287 MB |
| Find a text in all 300,000,000 cells | 1.9 s | 417 MB |

Only the parts on screen are read, so memory stays under half a gigabyte. Excel holds the whole sheet in memory.

**The slow part is the first open, and it stays slow.** Rule 0 makes Excel the only reader of a company file, so a file
seen for the first time is read by a hidden Excel. Excel itself must open a 300-million-cell workbook first, and that
takes minutes and several gigabytes. Later opens use the fast copy (0.1 s).

## Measured: the data store (questions, reports, dashboards)
- A 1M x 300 table is 1.9 GB on disk. Its header opens in 0.01 to 0.04 s, and 3 columns load in 0.03 s (20 MB).
- **Found and fixed on 2026-11-14:** every question loaded every column, and a report union loaded every period with
  every column. A question over a 3-month report of 1M x 300 per month ran **out of memory** on a 15 GB machine.
  Now the engine reads only the columns the question uses (filter, formulas, grouping, sorting, shown columns, join keys).
  The same question takes **1.5 s with 437 MB peak**. Report unions also merge text dictionaries instead of making one
  string per row. The provider now keeps 4 unions in memory instead of all of them (`TableProvider.MaxCached`).
- A question that shows every column (`SELECT * ... LIMIT 10`) filters and sorts on the columns it needs, then reads
  the 300 columns only for the rows it returns. Over MCP on 3M x 300, `describe_table` went from out of memory to 0.2 s.
  Still large: a Data-mode view with no column list and no limit returns every row of every column. The next step is
  result paging, so such a view loads only the rows on screen.

## How far from Excel 365
**Has:** open, scroll, select, edit with Excel's typing rules, about 170 calculated functions, fill series, copy, cut and
paste with moving references, undo and redo, number formats with a Format Cells dialog, bold, italic, colours, filter and
sort on many columns, find and replace, freeze panes, column widths, save through Excel, Excel's main shortcuts.
Data mode: folder watching and automatic loading, version history per period, quality checks, joins, saved views,
alerts, recipes with schedules, sandboxes, a library, dashboards with tiles, and rule-based questions.

**Missing, roughly by how often Excel users need it:**
1. Insert and delete rows, columns and sheets; merged cells; row heights; borders, alignment and wrap on screen
2. Charts (the plan allows hand-drawn SVG only, with no new packages)
3. Pivot tables (the store can compute them; there is no pivot interface)
4. Conditional formatting, data validation and drop-down lists
5. Dynamic arrays and spill (FILTER, SORT, UNIQUE, SEQUENCE), LET and LAMBDA
6. Named ranges, Excel tables (ListObjects) and structured references
7. Comments and notes, hyperlinks, pictures and shapes
8. Printing and page layout
9. Power Query and the data model (the Data mode covers part of this)
10. Macros and VBA, Office Scripts
11. Co-authoring, OneDrive and SharePoint (excluded by rule 0: data never leaves the laptop)
12. Copilot (excluded today: no AI model, see below)

Full parity with Excel 365 is years of work for a large team. The realistic goal is Excel's daily 20% done well and
much faster, plus the Data mode, which Excel does not have.

## Decisions only the owner (as IT manager) can make
1. **AI reading the data. Decided on 2026-11-14: AI may see everything, cell values included.** ExcelApp gets read-only agent access: a SQL-like `--query` command and an MCP server (`--mcp`), see `docs\AI-ACCESS.md`. The choices it was made from: Rule 0.1 says real data never leaves this laptop, and rule 0.10 says there is no AI model
   in the build. Any cloud AI, including Claude, ChatGPT and Copilot, would send cell values off the laptop. The options:
   - (a) A local, read-only query command for the store (SQL-like, for example
     `ExcelApp.exe --query "select Region, sum(Sales) from sales group by Region"`). An agent on the laptop could call
     it, but whatever the agent sends to a cloud model leaves the laptop.
   - (b) A local AI model on the laptop. That needs IT to provide one offline, plus disk space (C: has 3.5 GB free).
   - (c) A cloud AI that sees only names, shapes and counts, never cell values. The owner would approve exactly which
     fields it may see.
2. **The dashboard tab.** The Data mode already has dashboards made of tiles. A free designer that looks like a modern
   web app means drag-and-drop tiles, SVG charts and filters. Choose which charts and controls come first.
3. **The order of the missing Excel features** above.
