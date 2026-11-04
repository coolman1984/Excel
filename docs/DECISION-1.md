# Decision 1: how to build the fast Excel viewer

Date: 2026-11-04. All numbers were measured on this PC unless a source link is given. "Not measured" means I have no number.

**Short answer:** COM solves *reading* a protected file, and it solves *scrolling*. It does not solve the *first opening*: the wait for Excel to open the file stays. Rust cannot build on this PC today. C# can, with nothing to install. The biggest open point is the owner's company-rule question about a saved fast copy, because that decides how much the viewer is worth.

## COM ways

### How I tested
- A small real protected company file (about 70 KB): only timings were printed, never contents.
- Made-up workbooks of 2, 6 and 20 million cells (20 columns of numbers and short text, one sheet, no formulas, no colours). These test large sizes without touching company data.
- **Surprise:** every made-up workbook that Excel saved here also came out wrapped in NASCA DRM. So the big tests were on truly protected files. It also means protection seems to be added to *any* workbook Excel saves on this PC, not only to company originals.
- Excel was hidden, no alerts, read-only. Only the Excel I started was closed. Windows commit reached 86.4% at most, under the 90% stop. I did not look for a bigger real file.

**Limits of these numbers:** this PC (4 cores, 16 GB) is not a weak laptop. Each size was run once. Other work was running at times. Real files with formulas, colours and many sheets will probably open slower (not measured).

### The result in one table

| Made-up protected file | 2 million cells | 6 million cells | 20 million cells |
|---|---|---|---|
| File size | 18 MB | 55 MB | 184 MB |
| Excel opens the file | 4.7 s | 16.2 s | 43.3 s |
| Excel memory, file open (empty Excel: 162 MB) | 213 MB | 309 MB | 632 MB |
| **Read one screen** (40 rows, typical; worst) | **2.2 ms; 12 ms** | **2.6 ms; 15 ms** | **1.5 ms; 20 ms** |
| Read all, block by block (1,000 rows each) | 2.0 s | 5.9 s | 19.3 s |
| Read all into one big array | 2.1 s | 7.0 s | not run (too much memory) |
| Excel's own search, nothing found (worst case) | 1.1 s | 4.9 s | 9.0 s |

Also: Excel itself takes about 4 to 6 s to start. On the small real protected file, opening took about 1 s and one screen took 2 to 5 ms, which fits the table.

**What this shows:** the screen read stays at about 2 ms however big the file is. A smooth screen allows about 16 ms per picture, so reading is not the slow part. Opening is, at about 2 seconds per million cells for simple data.

### The four ways

| Way | Works on a protected file? | Speed | Memory | Verdict |
|---|---|---|---|---|
| **A. Whole used range into one array** | Yes | About 1 million cells per second; the program waits until it ends | My program grew about 40 bytes per cell (232 MB at 6 million). Excel rose by about 250 MB more during the read. At 20 million: over 2 GB in total (my estimate, not run) | Fine for small files. Wrong for huge files on a weak laptop |
| **B. Block by block** | Yes | Same total speed, about 1 million cells per second. 1,000-row blocks take about 20 ms each, so the screen can stay alive between blocks | My program stayed at 44 to 60 MB at every size | **Use for background work: search and sort** |
| **C. Excel saves a plain copy** | **No.** I tried an ordinary SaveAs once. The copy was still protected (it starts with the NASCA mark). I deleted it. I tried no other route, on purpose, because that would be working around the protection | SaveAs itself was fast on the small file, but a copy gives no speed gain, since only Excel can read it | Same as the file | **Dead end.** Also, a protected copy is no use to Calamine |
| **D. Only what is on the screen** | Yes | About 2 ms per screen, at any size | My program about 29 MB; Excel stays as in the table | **Use for scrolling** |

For comparison, reading the same screen one cell at a time took about 1,500 ms instead of 5 ms (measured in PowerShell). That must never be used.

**All ways keep the same two costs:** the wait for Excel to open the file, and Excel's memory (632 MB at 20 million cells). Nothing appears on screen until Excel finishes opening, so "load step by step" cannot apply to that first wait.

### The saved fast copy: possible, but allowed?

I built one on **made-up data only**, in my own simple format:

| | 2 million | 6 million | 20 million |
|---|---|---|---|
| Time to build it (Excel must be open) | 2.5 s | 8.1 s | 19.3 s |
| Size on disk | 21 MB | 64 MB | 214 MB |
| **Reopen it, no Excel at all** | 4 ms | 2 ms | **6 ms** instead of 43 s |
| One screen (file already in the PC's memory; a cold disk was not measured) | 0.06 ms | 0.08 ms | 0.08 ms |
| Full search, nothing found | 0.26 s | 0.77 s | 1.6 s instead of 9 s |
| My program's memory | | | about 56 MB, and no Excel |

The speed gain is huge. But I checked the file my program wrote: **it has no NASCA mark, so it is not protected.** Company data in that file would sit outside the protection. That is why this is a rule question and not a technical one. The project's company policy already says that being able to read a file gives no permission to keep an unprotected copy or store it in a database. **Until the owner has an answer, the viewer saves nothing to disk.** I did not make a fast copy of any real file.

### Other COM risks
- Excel handles one request at a time. While it is doing a long search, a screen read will probably have to wait (not measured). So search must run in small blocks, not as one long Excel search.
- A hidden pop-up from Excel can freeze the reader. Microsoft warns about this for programs running with no user logged on, and does not support Office automation in that case ([Microsoft](https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2)). The viewer runs in the owner's own session, but it still needs a separate worker with a time limit, as the company policy requires.
- After closing, my test Excel sometimes took more than 7 s to leave. It always left by itself.

## Lighter choices

### Building the program

| Tool | Needs to build | Lighter choice | What is lost |
|---|---|---|---|
| **Rust** (now 2.51 GB on C:) | **Cannot build here.** Even the smallest program failed: `linker link.exe not found`. It needs Microsoft's C++ build tools and the Windows SDK ([Rust docs](https://rust-lang.github.io/rustup/installation/windows-msvc.html)). Microsoft lists 2.3 GB to 60 GB of disk for Build Tools ([Microsoft](https://learn.microsoft.com/en-us/visualstudio/releases/2022/system-requirements)). C: has 3 GB free, and installs are not allowed now | **C# with the .NET 8 SDK, already installed.** Small test programs built in 7 to 12 s, with 130 to 150 MB at peak (first cold build: 35 s, 418 MB). Sampled every 0.3 s, so short peaks may be missed. No download | See "Is C# a better fit?" below |
| **Slint** | Cannot be measured without a Rust build. The licence is fine: free for a desktop app if the app shows an "About Slint" box or a public page shows a badge ([Slint](https://slint.dev/terms-and-conditions)) | Windows Forms, inside .NET | Modern look |
| **DuckDB in Rust** | From source it compiles C++, and its makers say build time rose "dramatically" ([duckdb-rs](https://github.com/altertable-ai/duckdb-rs)). Memory: not measured. A ready-made library can be downloaded instead ([DuckDB](https://duckdb.org/docs/current/clients/rust/troubleshoot)) | Leave it out now | See below |
| **.NET Framework 4.8, inside Windows** | Lightest of all: built in 1.8 s with the compiler already in Windows; 7 KB program | Fallback, for PCs without .NET | Old C# language; slower grid (below) |

### Running the program

| Tool | Memory or size when running | Lighter choice | What is lost |
|---|---|---|---|
| **Excel itself** (required for protected files) | 162 MB empty, 632 MB at 20 million cells | None for protected files. It can be closed when idle | The next opening waits again |
| **Slint** | Not measured. I found no fair number | **Windows Forms grid with 1 million rows on demand: 55 MB on .NET 8, 73 MB on .NET Framework 4.8.** The WPF grid, also tested: 144 MB | Modern look |
| Grid speed, random long jump plus a full redraw (a hard test, off screen; real scrolling not measured) | Windows Forms .NET 8: **64 ms**. .NET Framework 4.8: 118 ms. WPF: 247 ms | Windows Forms .NET 8 is best of the three. 64 ms is not fully smooth for jumps, so my own drawn grid may be needed later (not measured) | More code |
| **DuckDB** | Its makers say at least 125 MB per thread, and aim for 1 to 4 GB per thread ([DuckDB](https://duckdb.org/docs/2.0/guides/performance/environment)). The .NET package is a 105 MB download ([NuGet](https://www.nuget.org/packages/DuckDB.NET.Bindings.Full)) | **Leave it out.** It cannot read protected files. It would only store a saved copy, which may not be allowed. My simple copy searched 20 million cells in 1.6 s with 56 MB | Clever queries and big sorts on disk. Add it back only if a test shows the need |
| **Apache Arrow** | Not measured | Drop it; plain arrays | Nothing needed |
| **memmap2** | Not measured | Drop it. If a saved copy is allowed, .NET has this built in | Nothing needed |
| **Rayon** | Not measured | Drop it. .NET has this built in. COM cannot use it anyway, since Excel does one request at a time | Nothing needed |
| **SQLite** | Not measured | Drop it; a small settings file, with JSON built into .NET | Nothing needed |
| **Calamine** | Not measured | Not useful on protected files: it cannot read them. Keep the idea only for plain files (see question 2). The .NET equivalent to test later: Sylvan.Data.Excel 0.5.8, MIT licence, 280 KB ([NuGet](https://www.nuget.org/packages/Sylvan.Data.Excel)), not tested and needs a download | Plain files go through Excel too, and open slower (not measured) |

All packages named here were checked as real on 2026-11-04: calamine 0.38.0, slint 1.18.1, duckdb 1.20000.1, rayon 1.12.0, memmap2 0.9.12, arrow 60.1.0 (crates.io), and DuckDB.NET.Bindings.Full 1.5.6.1, Sylvan.Data.Excel 0.5.8 (NuGet).

### Is C# a better fit than Rust?
**Yes, for this job on this PC.** The reasons, in order:
1. **It builds today.** Rust does not, and its missing tools may not fit on the disk.
2. **Rust's speed gives little here.** On protected files the slow part is Excel, and a screen read takes about 2 ms from C#. Faster code in my program cannot make Excel open faster.
3. **Memory is decided by Excel.** The grid is 55 MB and Excel is up to 632 MB. A lighter grid saves little of that total.
4. **COM is simple in C#.** The test tool talks to Excel directly with a few lines.

**What is lost:** Rust's single small program with nothing else needed, and probably less memory for my program (not measured). Rust could be the better choice if most big files were plain and never touched Excel. That is question 2. Microsoft's way of making C# into one native program (Native AOT) does not help, because it has no built-in COM ([Microsoft](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/)).

**A date to know:** Microsoft support for .NET 8 ends on 2026-11-10; .NET 10 runs to 2028-11-15 ([Microsoft](https://learn.microsoft.com/en-us/lifecycle/products/microsoft-net-and-net-core)). This PC has the .NET 10 runtime but only the .NET 8 builder. I ran the .NET 8 grid test on .NET 10 and it worked at the same speed. So: build with what is here, run on .NET 10, and move the builder to .NET 10 when an install is allowed.

## Recommendation

**One design: a small C# Windows Forms viewer that uses Excel as its hidden reader, keeps nothing on disk, and has only the tools it needs.**

- **Opening:** a separate worker starts a hidden, read-only Excel of its own. The viewer shows an honest "Excel is opening the file" message. If the owner has a normal Excel open, the viewer does not touch it.
- **Scrolling:** Way D. Read only the visible rows, and keep the last few thousand rows in memory (5,000 rows of 20 columns is about 4 MB).
- **Search and sort:** Way B, in the background, in small blocks so the screen stays free. Sort reads one column and sorts row numbers in my program; the file does not change (not measured).
- **Repeat opening:** keep the hidden Excel open while the viewer is used, then close it after some idle time, so going back to the same file is instant.
- **Fit to the machine:** watch Windows commit. Use smaller blocks and memory when it is high, and warn before opening when it is near the limit.
- **Saved fast copy:** not built. It is one switch for later, only with a written company answer.
- **Left out:** Rust, Slint, Calamine, DuckDB, Arrow, memmap2, Rayon, SQLite. This differs from the lead's opinion, which kept Rust, Slint, Calamine and DuckDB.
- **Fallback:** .NET Framework 4.8 if company PCs have no newer .NET.

**The honest weak point.** If all big files are protected and no saved copy is allowed, the viewer opens files no faster than Excel does, because Excel does the opening. The viewer would then only give light, steady scrolling and searching after the wait. That may not be worth building. If a saved copy is allowed, or many files are plain, the gain is very large (6 ms instead of 43 s at 20 million cells). So the company answer and the first test decide whether to build at all.

## First test

**Goal: decide "build or do not build" on the largest real file, in about 10 minutes, before writing the app.**

1. The owner names the largest real protected file, normally used.
2. On that file, time a normal double-click opening in Excel, and note how scrolling and searching feel.
3. Run the tool now in this project, `tools\com-bench`. It starts its own hidden, read-only Excel, prints only timings and sizes, and refuses to make a fast copy of a real file:
   `dotnet build -c Release`, then `dotnet bench.dll bench "<file>" --nowhole`
4. Repeat on a truly weak laptop if one exists.

**How to read the result:**

| Check | Pass |
|---|---|
| One screen read | Typical 16 ms or less; worst 100 ms or less |
| Windows commit | Stays under 90%; Excel closes by itself |
| Opening time | The owner says if it is acceptable (the limit is the owner's) |
| Worth building? | Scrolling or searching in normal Excel is clearly a problem on that file, or a saved copy is allowed |

Results go in the private HERDR task notes, not in this repository, because they describe a company file.

If it passes, the next step is a very small viewer: one window, one sheet, Way D only, to feel real scrolling on screen.

## Questions for the owner

1. **The company-rule question (the most important): may a program keep a saved fast copy of a protected file's data on this PC?** This copy would not have NASCA protection on it. Choose one: (a) no, never; (b) only in memory, erased on closing; (c) yes on disk, with written approval from IT or security. Who in the company can answer? Until then, the answer is treated as "only in memory".
2. **Are most of your big files protected, or do some come as plain files, for example exports from a system?**
3. **Which file is the largest one, and may I test it as described above?**
4. **Does the company allow a self-built program?** On this PC only, or on others too? On this PC a self-built program ran fine and used Excel normally, but that does not prove it is allowed.
5. **Do you have a truly weak laptop for the test?**
6. **What do you do in these files: only look and search, or also sort and filter? And how long a first opening can you accept?** For example, one minute.
