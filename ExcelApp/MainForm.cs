// One window holding one WebView2. Native file dialogs, clipboard flags and the close prompt live here.
using System.Text.Json;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelApp;

public sealed class MainForm : Form
{
    public readonly WebView2 View = new();
    public readonly Session Session;
    public readonly Bridge Bridge;
    readonly ExcelApp.Agent.AgentLink agentLink;
    readonly string? devUrl; readonly string uiFolder;
    public bool Unattended; bool allowClose, closeAsked, closeAfterSave;
    public string? OpenOnStart;
    public Task<bool> Ready => ready.Task; readonly TaskCompletionSource<bool> ready = new();

    public MainForm(string? devUrl = null, string? uiFolder = null)
    {
        this.devUrl = devUrl;
        this.uiFolder = uiFolder ?? FindUi();
        Text = "ExcelApp"; ClientSize = new Size(1400, 860); MinimumSize = new Size(900, 600); AllowDrop = true;
        View.Dock = DockStyle.Fill; View.AllowExternalDrop = false; Controls.Add(View);
        Session = new Session();
        Session.Data = ExcelApp.Data.DataHost.Create((name, payload) => Session.Emit(name, payload)); // plan 4
        if (Session.Data.Services != null) Session.Journal = new EditJournal(Session.Data.Services.Store); // auto-save journal in the store
        Bridge = new Bridge(Session);
        Session.PickOpenPath = () => OnUi(() =>
        {
            using var d = new OpenFileDialog { Filter = "Excel files|*.xlsx;*.xlsm;*.xlsb;*.xls;*.csv|All files|*.*", Title = "Open" };
            return d.ShowDialog(this) == DialogResult.OK ? d.FileName : null;
        });
        Session.PickSavePath = src => OnUi(() =>
        {
            string ext = Path.GetExtension(src).ToLowerInvariant();
            using var d = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel Binary Workbook (*.xlsb)|*.xlsb",
                FilterIndex = ext == ".xlsm" ? 2 : ext == ".xlsb" ? 3 : 1, InitialDirectory = Path.GetDirectoryName(src), FileName = Path.GetFileNameWithoutExtension(src), AddExtension = true, Title = "Save As",
            };
            return d.ShowDialog(this) == DialogResult.OK ? d.FileName : null;
        });
        Session.ClipboardSet = text => OnUi<object?>(() => { ClipboardSink(text); return null; });
        Session.ClipboardGet = () => OnUi(() => Clipboard.ContainsText() ? Clipboard.GetText() : null);
        Session.TitleChanged = (name, dirty) => { try { BeginInvoke(() => Text = name.Length == 0 ? "ExcelApp" : "ExcelApp - " + name + (dirty ? " *" : "")); } catch (InvalidOperationException) { } };
        Session.CloseAnswer = OnCloseAnswer;
        Bridge.Post = json => { try { BeginInvoke(() => { if (View.CoreWebView2 != null) View.CoreWebView2.PostWebMessageAsJson(json); }); } catch (InvalidOperationException) { } };
        DragEnter += (s, e) => e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true ? DragDropEffects.Copy : DragDropEffects.None;
        DragDrop += (s, e) => { if (e.Data?.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0) { try { Session.Open(f[0]); } catch (UserError) { } } };
        FormClosing += OnClosing;
        Load += async (s, e) => await Init();
        // AI agents (owner's decision, 2026-11-14) reach the open workbook through ExcelApp.exe --mcp and this door.
        agentLink = new ExcelApp.Agent.AgentLink(Bridge.Dispatch, cmd => Bridge.Notify("agentChanged", new { Cmd = cmd }));
        agentLink.Start();
        FormClosed += (s, e) => agentLink.Dispose();
    }

    // The built interface next to the program; else the temporary stub of the source tree.
    static string FindUi()
    {
        string built = Path.Combine(AppContext.BaseDirectory, "ui");
        if (File.Exists(Path.Combine(built, "index.html"))) return built;
        for (var d = new DirectoryInfo(AppContext.BaseDirectory); d != null; d = d.Parent)
        {
            string stub = Path.Combine(d.FullName, "ExcelApp", "ui-stub");
            if (File.Exists(Path.Combine(stub, "index.html"))) return stub;
        }
        return built;
    }

    T OnUi<T>(Func<T> f) => InvokeRequired ? (T)Invoke(f)! : f();

    /// <summary>Puts text on this PC's clipboard, and asks Windows to keep it out of clipboard history and cloud sync.</summary>
    public static Action<string> ClipboardSink = text =>
    {
        var data = new DataObject();
        data.SetData(DataFormats.UnicodeText, text); data.SetData(DataFormats.Text, text);
        data.SetData("ExcludeClipboardContentFromMonitorProcessing", new MemoryStream(BitConverter.GetBytes(1)));
        data.SetData("CanIncludeInClipboardHistory", new MemoryStream(BitConverter.GetBytes(0)));
        data.SetData("CanUploadToCloudClipboard", new MemoryStream(BitConverter.GetBytes(0)));
        Clipboard.SetDataObject(data, true, 5, 100);
    };

    async Task Init()
    {
        var initTimer = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            string data = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelApp", "WebView2");
            CoreWebView2Environment env;
            try { env = await CoreWebView2Environment.CreateAsync(null, data); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException)
            { Log.Write("webview data folder fallback"); env = await CoreWebView2Environment.CreateAsync(null, Path.Combine(Path.GetTempPath(), "ExcelAppWebView2")); }
            if (Unattended) Console.WriteLine($"START environment {initTimer.ElapsedMilliseconds} ms");
            await View.EnsureCoreWebView2Async(env);
            if (Unattended) Console.WriteLine($"START webview {initTimer.ElapsedMilliseconds} ms");
        }
        catch (Exception e)
        {
            Log.Write("webview failed " + e.GetType().Name);
            if (Unattended) { Console.WriteLine("ERROR webview " + e.GetType().Name); Environment.Exit(7); }
            MessageBox.Show("ExcelApp needs the Microsoft Edge WebView2 runtime, and it could not be started.", "ExcelApp");
            Environment.Exit(7);
        }
        var w = View.CoreWebView2; var st = w.Settings;
        st.AreDefaultContextMenusEnabled = false; st.AreBrowserAcceleratorKeysEnabled = false; st.IsZoomControlEnabled = false;
        st.IsStatusBarEnabled = false; st.IsSwipeNavigationEnabled = false; st.AreDevToolsEnabled = devUrl != null;
        w.NewWindowRequested += (s, e) => e.Handled = true;
        bool Trusted(string source) => Uri.TryCreate(source, UriKind.Absolute, out var u) &&
            (devUrl == null ? u.Scheme == "https" && u.Host == "app.excel" : new Uri(devUrl).GetLeftPart(UriPartial.Authority) == u.GetLeftPart(UriPartial.Authority));
        w.NavigationStarting += (s, e) => { if (!Trusted(e.Uri)) e.Cancel = true; };
        w.PermissionRequested += (s, e) => e.State = CoreWebView2PermissionState.Deny;
        w.DownloadStarting += (s, e) => e.Cancel = true;
        w.WebMessageReceived += (s, e) => { try { if (!Trusted(e.Source)) return; _ = Bridge.HandleMessage(e.WebMessageAsJson is { Length: > 0 } j && j.StartsWith('"') ? JsonSerializer.Deserialize<string>(j)! : e.WebMessageAsJson); } catch (Exception ex) { Log.Write("message " + ex.GetType().Name); } };
        if (devUrl != null) w.Navigate(devUrl);
        else
        {
            if (!Directory.Exists(uiFolder)) { Log.Write("ui folder missing"); }
            w.SetVirtualHostNameToFolderMapping("app.excel", uiFolder, CoreWebView2HostResourceAccessKind.Allow);
            w.Navigate("https://app.excel/index.html");
        }
        w.NavigationCompleted += (s, e) => { if (Unattended) Console.WriteLine($"START navigation {initTimer.ElapsedMilliseconds} ms"); ready.TrySetResult(e.IsSuccess); };
        // The file named on the command line opens once the interface has asked for "ready", so no event is missed.
        Session.UiReady = () => { var p = OpenOnStart; OpenOnStart = null; if (p != null) try { Session.Open(p); } catch (UserError) { } };
    }

    public void PlaceOffScreen() { StartPosition = FormStartPosition.Manual; Location = new Point(-32000, -32000); ShowInTaskbar = false; }

    // ---- closing ------------------------------------------------------------------------------
    void OnClosing(object? s, FormClosingEventArgs e)
    {
        if (Session.ChildTimedOut) { Session.CancelBuildForClose(); return; }
        if (allowClose || Unattended) { Session.CancelBuildForClose(); return; }
        if (Session.Saving) { e.Cancel = true; if (!closeAfterSave) { closeAfterSave = true; Bridge.Post?.Invoke("{\"event\":\"closing\",\"phase\":\"Finishing the save...\"}"); WaitSaveThenClose(); } return; }
        if (Session.Dirty) { e.Cancel = true; if (!closeAsked) { closeAsked = true; Bridge.Post?.Invoke(JsonSerializer.Serialize(new { @event = "closeRequested", fileName = Session.FileName })); } return; }
        Session.CancelBuildForClose();
    }

    void WaitSaveThenClose() => Task.Run(async () => { while (Session.Saving) await Task.Delay(200); BeginInvoke(() => { allowClose = !Session.Dirty; closeAfterSave = false; Close(); }); });

    void OnCloseAnswer(string action)
    {
        closeAsked = false;
        if (action == "dontsave") BeginInvoke(() => { allowClose = true; Close(); });
        else if (action == "save")
            Task.Run(async () =>
            {
                try { await Bridge.Dispatch("{\"id\":0,\"cmd\":\"save\"}"); } catch (Exception) { }
                BeginInvoke(() => { if (!Session.Dirty) { allowClose = true; Close(); } });
            });
    }
}
