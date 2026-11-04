// One window: pick a file, see it, scroll it, search it, sort it.
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using ExcelViewer.Core;

namespace ExcelViewer;

public sealed class MainForm : Form
{
    readonly ToolStrip bar = new() { GripStyle = ToolStripGripStyle.Hidden };
    readonly ToolStripButton openBtn = new("Open file...");
    readonly ToolStripComboBox sheetBox = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
    readonly ToolStripButton headerBtn = new("First row is header") { CheckOnClick = true };
    readonly ToolStripTextBox searchBox = new() { Width = 220 };
    readonly ToolStripComboBox scopeBox = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 110 };
    readonly ToolStripButton searchBtn = new("Search");
    readonly ToolStripButton cancelBtn = new("Stop") { Enabled = false };
    readonly ToolStripButton rebuildBtn = new("Read file again");
    readonly DataGridView grid = new();
    readonly ContextMenuStrip gridMenu = new();
    readonly ToolStripMenuItem copyItem = new("Copy");
    readonly ListView hitsList = new();
    readonly SplitContainer split = new() { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal };
    readonly StatusStrip status = new();
    readonly ToolStripStatusLabel sizeLabel = new(), infoLabel = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft }, warnLabel = new() { ForeColor = Color.DarkRed };

    readonly Panel busy = new() { Dock = DockStyle.Fill, BackColor = SystemColors.Window, Visible = false };
    readonly Label busyTitle = new() { AutoSize = false, Height = 44, Dock = DockStyle.Top, Font = new Font("Segoe UI", 14f), TextAlign = ContentAlignment.BottomCenter };
    readonly Label busyNote = new() { AutoSize = false, Height = 70, Dock = DockStyle.Top, TextAlign = ContentAlignment.MiddleCenter };
    readonly ProgressBar busyBar = new() { Dock = DockStyle.Top, Height = 16, Style = ProgressBarStyle.Marquee };
    readonly Button busyCancel = new() { Text = "Cancel", Width = 100, Height = 30 };
    readonly System.Windows.Forms.Timer tick = new() { Interval = 500 };
    readonly Stopwatch busyClock = new();

    string sourcePath; CacheData cache; SheetReader view; int sheetIndex = -1; int firstDataRow; int[] rowMap;
    int sortCol = -1; bool sortDesc; CancellationTokenSource sortCts, searchCts; Process buildProc; string busyPhase = "";
    public bool Unattended; // the timing and picture runs: never wait on a box nobody can see
    void Tell(string msg) { if (Unattended) Console.WriteLine("MESSAGE: " + msg); else MessageBox.Show(this, msg, "Excel Viewer"); }
    readonly List<Hit> hits = new(); readonly object hitsLock = new(); int hitCol = -1;
    bool searchDone; string searchNote = "";

    public MainForm()
    {
        Text = "Excel Viewer"; Size = new Size(1280, 800); Font = new Font("Segoe UI", 9f); AllowDrop = true; KeyPreview = true;
        scopeBox.Items.AddRange(new object[] { "This sheet", "All sheets" }); scopeBox.SelectedIndex = 0;
        bar.Items.AddRange(new ToolStripItem[] { openBtn, new ToolStripSeparator(), new ToolStripLabel("Sheet:"), sheetBox, headerBtn, new ToolStripSeparator(), new ToolStripLabel("Find:"), searchBox, scopeBox, searchBtn, cancelBtn, new ToolStripSeparator(), rebuildBtn });

        grid.Dock = DockStyle.Fill; grid.VirtualMode = true; grid.ReadOnly = true; grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.AllowUserToOrderColumns = false; grid.AllowUserToResizeRows = false;
        grid.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect; grid.MultiSelect = true; grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing; grid.ColumnHeadersHeight = 26; grid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing; grid.RowHeadersWidth = 70;
        grid.BackgroundColor = SystemColors.Window; grid.BorderStyle = BorderStyle.None; grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable; // our own copy below takes the shown text grid.StandardTab = true;
        grid.RowTemplate.Height = 22; grid.EnableHeadersVisualStyles = true; grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
        typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(grid, true);
        grid.CellValueNeeded += (s, e) => { if (view != null) e.Value = view.Cell(MapRow(e.RowIndex), e.ColumnIndex).Text; };
        grid.CellFormatting += GridCellFormatting;
        grid.RowPostPaint += GridRowPostPaint;
        grid.ContextMenuStrip = gridMenu; gridMenu.Items.Add(copyItem); copyItem.ShortcutKeyDisplayString = "Ctrl+C"; copyItem.Click += (s, e) => CopySelection();
        grid.CellMouseDown += (s, e) =>
        {
            // A right-click on a cell outside the selection selects that cell first.
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0 && !grid[e.ColumnIndex, e.RowIndex].Selected) { grid.ClearSelection(); grid.CurrentCell = grid[e.ColumnIndex, e.RowIndex]; }
        };
        grid.ColumnHeaderMouseClick += (s, e) => { if (e.Button == MouseButtons.Left) ClickSort(e.ColumnIndex); };

        hitsList.Dock = DockStyle.Fill; hitsList.View = View.Details; hitsList.FullRowSelect = true; hitsList.VirtualMode = true; hitsList.MultiSelect = false; hitsList.HideSelection = false;
        hitsList.Columns.Add("Sheet", 200); hitsList.Columns.Add("Cell", 90); hitsList.Columns.Add("Shown", 800);
        hitsList.RetrieveVirtualItem += HitsRetrieve; hitsList.SelectedIndexChanged += (s, e) => { if (hitsList.SelectedIndices.Count > 0) JumpToHit(hitsList.SelectedIndices[0]); };

        split.Panel1.Controls.Add(grid); split.Panel2.Controls.Add(hitsList); split.Panel2Collapsed = true;
        status.Items.AddRange(new ToolStripItem[] { sizeLabel, infoLabel, warnLabel });

        var centre = new Panel { Width = 560, Height = 200 };
        busyBar.Parent = centre; busyNote.Parent = centre; busyTitle.Parent = centre; // docked last-first
        busyCancel.Parent = centre; busyCancel.Location = new Point(230, 150);
        busy.Controls.Add(centre); busy.Resize += (s, e) => centre.Location = new Point(Math.Max(0, (busy.Width - centre.Width) / 2), Math.Max(0, (busy.Height - centre.Height) / 2));

        Controls.Add(split); Controls.Add(busy); Controls.Add(status); Controls.Add(bar);
        busy.BringToFront();

        openBtn.Click += (s, e) => PickFile();
        sheetBox.SelectedIndexChanged += (s, e) => { if (sheetBox.SelectedIndex >= 0 && sheetBox.SelectedIndex != sheetIndex) ShowSheet(sheetBox.SelectedIndex); };
        headerBtn.Click += (s, e) => { firstDataRow = headerBtn.Checked ? 1 : 0; ResetSort(); RefreshGrid(); };
        searchBtn.Click += (s, e) => StartSearch();
        searchBox.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; StartSearch(); } };
        cancelBtn.Click += (s, e) => searchCts?.Cancel();
        rebuildBtn.Click += (s, e) => { if (sourcePath != null) { CloseFile(); BuildThenOpen(sourcePath, "You asked to read the file again."); } };
        busyCancel.Click += (s, e) => CancelBuild();
        tick.Tick += (s, e) => OnTick();
        DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
        DragDrop += (s, e) => { if (e.Data.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0) OpenFile(f[0]); };
        FormClosing += (s, e) => { CancelBuild(); searchCts?.Cancel(); sortCts?.Cancel(); };
        FormClosed += (s, e) => { view?.Dispose(); view = null; };
        SetEmpty();
    }

    // ---------- Opening ----------

    void SetEmpty() { sizeLabel.Text = ""; infoLabel.Text = "Pick an Excel file with \"Open file...\", or drop one on this window."; warnLabel.Text = ""; rebuildBtn.Enabled = false; }

    void PickFile()
    {
        using var d = new OpenFileDialog { Filter = "Excel files|*.xlsx;*.xlsm;*.xlsb;*.xls|All files|*.*", Title = "Pick an Excel file" };
        if (d.ShowDialog(this) == DialogResult.OK) OpenFile(d.FileName);
    }

    public void OpenFile(string path)
    {
        CloseFile(); sourcePath = path;
        if (!CachePolicy.ThisPcApproved) { Tell("The fast copy is approved for one laptop only, and this is not that laptop. Nothing was read."); SetEmpty(); return; }
        string cp = CachePolicy.PathFor(path);
        if (CacheData.HeaderMatches(cp, path))
        {
            try { LoadCache(cp); return; }
            catch (InvalidDataException) { CloseFile(); sourcePath = path; BuildThenOpen(path, "The saved copy could not be read, so the file is being read again."); return; }
        }
        BuildThenOpen(path, File.Exists(cp) ? "The file has changed since the last read." : "This is the first time this file is read.");
    }

    void LoadCache(string cachePath)
    {
        cache = CacheData.Load(cachePath);
        Text = "Excel Viewer - " + Path.GetFileName(sourcePath);
        sheetBox.Items.Clear();
        foreach (var d in cache.Sheets) sheetBox.Items.Add(d.Name + (d.Hidden ? " (hidden)" : ""));
        int best = 0; long most = -1;
        for (int i = 0; i < cache.Sheets.Count && most < 0; i++) if (!cache.Sheets[i].Hidden && cache.Sheets[i].Rows > 0) { best = i; most = 0; }
        rebuildBtn.Enabled = true;
        ShowSheet(best); sheetBox.SelectedIndex = best;
    }

    public void ShowSheet(int i)
    {
        view?.Dispose(); sheetIndex = i; var d = cache.Sheets[i]; view = new SheetReader(cache, d);
        sortCts?.Cancel(); sortCol = -1; rowMap = null; hitCol = -1;
        grid.SuspendLayout(); grid.RowCount = 0; grid.Columns.Clear();
        bool header = LooksLikeHeader(d); headerBtn.Checked = header; firstDataRow = header ? 1 : 0;
        var cols = new DataGridViewColumn[d.Cols];
        for (int k = 0; k < d.Cols; k++)
            cols[k] = new DataGridViewTextBoxColumn { SortMode = DataGridViewColumnSortMode.Programmatic, FillWeight = 1, Width = Math.Clamp((int)(d.ColWidths[k] * 7) + 8, 50, 500), HeaderText = ColumnTitle(k) };
        grid.Columns.AddRange(cols);
        grid.ResumeLayout();
        RefreshGrid();
        sizeLabel.Text = $"{d.Rows:N0} rows x {d.Cols:N0} columns"; infoLabel.Text = "Read from the fast copy. Click a column title to sort.";
        int warn = view.ColumnsWithFormatWarnings();
        warnLabel.Text = warn == 0 ? "" : $"⚠ {warn} column(s) have a number format that could not be matched to Excel: those cells show the plain value, marked ⚠";
    }

    bool LooksLikeHeader(SheetInfo d)
    {
        if (d.Rows < 2 || d.Cols == 0) return false;
        int text = 0; for (int k = 0; k < d.Cols; k++) { byte kd = view.Cell(0, k).Kind; if (kd == CellKind.Number) return false; if (kd == CellKind.Text) text++; }
        return text > 0;
    }

    string ColumnTitle(int k)
    {
        string letter = CacheBuilder.Letters(view.Sheet.FirstCol + k) + (view.Sheet.ColWidths[k] == 0 ? " (hidden)" : "");
        if (firstDataRow == 0 && !headerBtn.Checked) return letter;
        string t = view.Cell(0, k).Text;
        return string.IsNullOrEmpty(t) ? letter : letter + "  " + t;
    }

    void RefreshGrid()
    {
        if (view == null) return;
        var d = view.Sheet;
        for (int k = 0; k < grid.Columns.Count; k++) { grid.Columns[k].HeaderText = ColumnTitle(k); grid.Columns[k].HeaderCell.SortGlyphDirection = k == sortCol ? (sortDesc ? SortOrder.Descending : SortOrder.Ascending) : SortOrder.None; }
        grid.RowCount = Math.Max(0, d.Rows - firstDataRow);
        grid.Invalidate();
    }

    int MapRow(int gridRow) => rowMap != null ? rowMap[gridRow] : gridRow + firstDataRow;

    void GridCellFormatting(object s, DataGridViewCellFormattingEventArgs e)
    {
        if (view == null || e.RowIndex < 0) return;
        var c = view.Cell(MapRow(e.RowIndex), e.ColumnIndex);
        if (c.Kind == CellKind.Number) e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        else if (c.Kind == CellKind.Bool || c.Kind == CellKind.Error) e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        if (c.Color != null) e.CellStyle.ForeColor = c.Color.Value;
        if (c.Warn) e.CellStyle.ForeColor = Color.DarkRed;
        if (e.ColumnIndex == hitCol && grid.CurrentCell != null && e.RowIndex == grid.CurrentCell.RowIndex) { e.CellStyle.BackColor = Color.Gold; e.CellStyle.SelectionBackColor = Color.Gold; e.CellStyle.SelectionForeColor = Color.Black; }
    }

    void GridRowPostPaint(object s, DataGridViewRowPostPaintEventArgs e)
    {
        if (view == null) return;
        string n = (view.Sheet.FirstRow + MapRow(e.RowIndex)).ToString();
        TextRenderer.DrawText(e.Graphics, n, grid.Font, new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth - 4, e.RowBounds.Height), SystemColors.ControlText, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
    }

    void CloseFile()
    {
        searchCts?.Cancel(); sortCts?.Cancel(); view?.Dispose(); view = null; cache = null; sheetIndex = -1; rowMap = null; sortCol = -1;
        grid.RowCount = 0; grid.Columns.Clear(); sheetBox.Items.Clear(); lock (hitsLock) hits.Clear(); hitsList.VirtualListSize = 0; split.Panel2Collapsed = true;
        SetEmpty(); Text = "Excel Viewer";
    }

    // ---------- Reading through the separate reader, which owns the hidden Excel ----------

    void BuildThenOpen(string path, string why)
    {
        double commit = Machine.Commit();
        if (commit >= Machine.StopAt) { Tell($"Windows memory is at {commit:F0}%. Reading a big file now could hang the PC, so nothing was started. Close some programs and try again."); SetEmpty(); return; }
        ShowBusy("Excel is opening the file...", why + " Only Excel can read a protected file, and big files can take about 2 minutes. Excel cannot show progress during this step." + (commit >= 85 ? $"\nWindows memory is already at {commit:F0}%. The read stops by itself at {Machine.StopAt:F0}%." : ""));
        var psi = new ProcessStartInfo(Environment.ProcessPath) { RedirectStandardOutput = true, RedirectStandardInput = true, UseShellExecute = false, CreateNoWindow = true };
        psi.ArgumentList.Add("--build"); psi.ArgumentList.Add(path);
        buildProc = Process.Start(psi);
        var proc = buildProc;
        proc.OutputDataReceived += (s, e) => { if (e.Data != null) BeginInvoke(() => OnBuildLine(e.Data)); };
        proc.EnableRaisingEvents = true;
        proc.Exited += (s, e) => { try { BeginInvoke(() => OnBuildExit(proc, path)); } catch (InvalidOperationException) { } };
        proc.BeginOutputReadLine();
    }

    void OnBuildLine(string line)
    {
        if (line == "OPENED" || line == "FORMATS") { busyPhase = "Reading number formats..."; }
        else if (line.StartsWith("READ ")) { busyBar.Style = ProgressBarStyle.Continuous; busyBar.Value = Math.Clamp(int.Parse(line[5..]), 0, 100); busyPhase = "Making the fast copy..."; }
        else if (line == "CHECKING") { busyPhase = "Checking dates and money against Excel..."; busyBar.Value = 100; }
        else if (line.StartsWith("RESULT ")) { lastResult = line; }
    }
    string lastResult = "";

    void OnBuildExit(Process proc, string path)
    {
        if (proc != buildProc) return;
        int code = proc.ExitCode; buildProc = null; HideBusy();
        if (code == 0)
        {
            try { LoadCache(CachePolicy.PathFor(path)); }
            catch (InvalidDataException) { SetEmpty(); infoLabel.Text = "The new copy could not be read."; Tell("The new copy could not be read. Nothing is shown."); }
            return;
        }
        string msg = code switch
        {
            2 => "Cancelled. Excel was closed and no copy was kept.",
            90 => $"Stopped because Windows memory reached {Machine.StopAt:F0}%. Excel was closed and no copy was kept. Close some programs and try again.",
            5 => "There is not enough free disk space for the copy. Nothing was kept.",
            _ => "Excel could not read this file. Nothing was kept.",
        };
        SetEmpty(); infoLabel.Text = msg; Tell(msg);
    }

    void CancelBuild()
    {
        if (buildProc == null || buildProc.HasExited) return;
        busyPhase = "Stopping. Excel closes as soon as it answers..."; busyCancel.Enabled = false;
        try { buildProc.StandardInput.Close(); } catch (IOException) { } catch (InvalidOperationException) { } // the reader then closes its own Excel
    }

    void ShowBusy(string title, string note) { busyTitle.Text = title; busyNote.Text = note; busyPhase = ""; busyBar.Style = ProgressBarStyle.Marquee; busyCancel.Enabled = true; busy.Visible = true; busy.BringToFront(); bar.Enabled = false; busyClock.Restart(); tick.Start(); OnTick(); }
    void HideBusy() { tick.Stop(); busy.Visible = false; bar.Enabled = true; busyClock.Stop(); }

    void OnTick()
    {
        if (busy.Visible)
        {
            if (busyPhase.Length > 0) busyTitle.Text = busyPhase;
            infoLabel.Text = $"Waiting {busyClock.Elapsed:m\\:ss}. Windows memory {Machine.Commit():F0}%.";
            if (busyClock.Elapsed.TotalMinutes > 10 && busyPhase.Length == 0) infoLabel.Text += " Excel may be waiting for something. It is never force-closed; check Task Manager if this goes on.";
        }
        if (searchCts != null) UpdateHits();
    }

    // ---------- Sort ----------

    public void ClickSort(int col)
    {
        if (view == null) return;
        bool desc = false; bool clear = false;
        if (col == sortCol) { if (!sortDesc) desc = true; else clear = true; }
        sortCts?.Cancel();
        if (clear) { ResetSort(); RefreshGrid(); infoLabel.Text = "Back to the file's own order."; return; }
        var cts = sortCts = new CancellationTokenSource(); var c = cache; var d = view.Sheet; int first = firstDataRow; var sw = Stopwatch.StartNew();
        infoLabel.Text = "Sorting...";
        Task.Run(() =>
        {
            try { var map = SortEngine.Sort(c, d, col, desc, first, cts.Token); BeginInvoke(() => { if (cts != sortCts || view?.Sheet != d) return; rowMap = map; sortCol = col; sortDesc = desc; RefreshGrid(); infoLabel.Text = $"Sorted in {sw.Elapsed.TotalSeconds:F2} s. The file is not changed."; }); }
            catch (OperationCanceledException) { }
        });
    }

    void ResetSort() { sortCts?.Cancel(); rowMap = null; sortCol = -1; }

    // ---------- Search ----------

    void StartSearch()
    {
        if (cache == null) return;
        string q = searchBox.Text; if (q.Length == 0) return;
        searchCts?.Cancel(); var cts = searchCts = new CancellationTokenSource();
        lock (hitsLock) hits.Clear(); hitsList.VirtualListSize = 0; split.Panel2Collapsed = false; split.SplitterDistance = Math.Max(100, split.Height - 200);
        var c = cache; var sheets = scopeBox.SelectedIndex == 1 ? Enumerable.Range(0, c.Sheets.Count).ToList() : new List<int> { sheetIndex };
        cancelBtn.Enabled = true; searchDone = false; var sw = Stopwatch.StartNew(); int pct = 0; tick.Start();
        searchNote = "Searching...";
        Task.Run(() =>
        {
            try
            {
                bool all = SearchEngine.Run(c, sheets, q, h => { lock (hitsLock) hits.Add(h); }, p => { pct = p; searchNote = $"Searching... {p}%"; }, cts.Token);
                BeginInvoke(() => { if (cts != searchCts) return; searchNote = all ? $"Found {hits.Count:N0} in {sw.Elapsed.TotalSeconds:F2} s." : $"Stopped at the first {SearchEngine.MaxHits:N0} matches ({sw.Elapsed.TotalSeconds:F2} s)."; EndSearch(); });
            }
            catch (OperationCanceledException) { BeginInvoke(() => { if (cts != searchCts) return; searchNote = "Search stopped."; EndSearch(); }); }
        });
    }

    void EndSearch() { cancelBtn.Enabled = false; searchDone = true; UpdateHits(); if (!busy.Visible) tick.Stop(); }

    void UpdateHits()
    {
        int n; lock (hitsLock) n = hits.Count;
        if (hitsList.VirtualListSize != n) hitsList.VirtualListSize = n;
        if (!busy.Visible) infoLabel.Text = searchDone || searchNote.Length > 0 ? searchNote + (searchDone ? "" : $" {n:N0} so far.") : infoLabel.Text;
    }

    readonly Dictionary<int, SheetReader> hitReaders = new();

    void HitsRetrieve(object s, RetrieveVirtualItemEventArgs e)
    {
        Hit h; lock (hitsLock) { if (e.ItemIndex >= hits.Count) { e.Item = new ListViewItem(new[] { "", "", "" }); return; } h = hits[e.ItemIndex]; }
        var d = cache.Sheets[h.Sheet];
        if (h.Sheet == sheetIndex && view != null) e.Item = Row(d, h, view);
        else { using var r = new SheetReader(cache, d); e.Item = Row(d, h, r); }
    }

    static ListViewItem Row(SheetInfo d, Hit h, SheetReader r) => new(new[] { d.Name, CacheBuilder.Letters(d.FirstCol + h.Col) + (d.FirstRow + h.Row), r.Cell(h.Row, h.Col).Text });

    void JumpToHit(int index)
    {
        Hit h; lock (hitsLock) { if (index >= hits.Count) return; h = hits[index]; }
        if (h.Sheet != sheetIndex) { sheetBox.SelectedIndex = h.Sheet; }
        int gridRow = rowMap != null ? Array.IndexOf(rowMap, h.Row) : h.Row - firstDataRow;
        if (gridRow < 0) { infoLabel.Text = "That match is in the header row, shown in the column title."; return; }
        hitCol = h.Col;
        grid.CurrentCell = grid[h.Col, gridRow];
        grid.FirstDisplayedScrollingRowIndex = Math.Max(0, gridRow - 5);
        grid.Invalidate();
    }

    // ---------- Copy ----------

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData) =>
        (grid.Focused && HandleGridKeys(keyData)) || base.ProcessCmdKey(ref msg, keyData);

    public bool HandleGridKeys(Keys keys)
    {
        if (keys is (Keys.Control | Keys.C) or (Keys.Control | Keys.Insert)) { CopySelection(); return true; }
        if (keys == (Keys.Control | Keys.A)) { infoLabel.Text = $"Select-all is off: it would be too heavy on a big sheet. Copy works up to {CopyText.MaxCells:N0} cells."; return true; }
        return false;
    }

    /// <summary>The shown text of the selected cells, in the order on screen. Null if there are too many.</summary>
    public string SelectionText()
    {
        if (view == null) return "";
        int n = grid.GetCellCount(DataGridViewElementStates.Selected);
        if (n == 0) return "";
        if (n > CopyText.MaxCells) return null;
        var cells = new List<(int, int)>(n);
        foreach (DataGridViewCell c in grid.SelectedCells) cells.Add((c.RowIndex, c.ColumnIndex));
        return CopyText.Build(cells, (row, col) => view.Cell(MapRow(row), col).Text);
    }

    public void CopySelection()
    {
        string text = SelectionText();
        if (text == null) { infoLabel.Text = $"Too many cells selected. Copy works up to {CopyText.MaxCells:N0} cells."; return; }
        if (text.Length == 0) { infoLabel.Text = "Nothing is selected."; return; }
        try { ClipboardSink(text); infoLabel.Text = $"Copied {grid.GetCellCount(DataGridViewElementStates.Selected):N0} cell(s)."; }
        catch (System.Runtime.InteropServices.ExternalException) { infoLabel.Text = "Windows would not let the copy go to the clipboard. Try again."; }
    }

    /// <summary>Puts text on this PC's clipboard, and asks Windows to keep it out of clipboard history and cloud sync.</summary>
    public Action<string> ClipboardSink = text =>
    {
        var data = new DataObject();
        data.SetData(DataFormats.UnicodeText, text); data.SetData(DataFormats.Text, text);
        data.SetData("ExcludeClipboardContentFromMonitorProcessing", new MemoryStream(BitConverter.GetBytes(1)));
        data.SetData("CanIncludeInClipboardHistory", new MemoryStream(BitConverter.GetBytes(0)));
        data.SetData("CanUploadToCloudClipboard", new MemoryStream(BitConverter.GetBytes(0)));
        Clipboard.SetDataObject(data, true, 5, 100);
    };

    public ToolStripMenuItem CopyMenuItem => copyItem;

    // ---------- For the timing run and the tests ----------

    public DataGridView Grid => grid;
    public SheetReader Reader => view;
    public CacheData Cache => cache;
    public int[] RowMap { get => rowMap; set => rowMap = value; }
    public int FirstDataRow => firstDataRow;
    public void Redraw() => RefreshGrid();
    public void PlaceOffScreen() { StartPosition = FormStartPosition.Manual; Location = new Point(-4000, -4000); ShowInTaskbar = false; }
    public void Source(string path) => sourcePath = path;
    public void LoadForTest(string cachePath) => LoadCache(cachePath);
    public int SortedColumn => sortCol;
    public bool SearchFinished => searchDone;
    public string InfoText => infoLabel.Text;
    public void SearchFor(string q) { searchBox.Text = q; StartSearch(); }
    public int HitCount { get { lock (hitsLock) return hits.Count; } }
    public void JumpToFirstHit() { if (HitCount > 0) { hitsList.SelectedIndices.Clear(); hitsList.SelectedIndices.Add(0); JumpToHit(0); } }
}
