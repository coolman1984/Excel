using ExcelApp.Data;
public static class StoreTests
{
    public static void Run(Action<string, bool, string> check, bool offline)
    {
        Review(check);
        var root = StoreRoot.ForTest(@"D:\ExcelAppTest\syn-plan4-store");
        var table = new Table("合成", new Column[] { new NumberColumn("n", new[] {1.5, 0.0}, new[] {false, true}), new NumberColumn("d", new[] {45000.0, 0.0}, new[] {false, true}, ColType.Date), TextColumn.From("t", new string?[] {"مرحبا", null}), new BoolColumn("b", new sbyte[] {1,-1}) });
        TableFile.Write(root, "round.xvt", table, new Dictionary<string,string> { ["test"] = "yes" });
        var read = TableFile.Open(root, "round.xvt");
        check("table round trip", read.Load().Name == table.Name && read.Load().Rows == 2 && read.Load().Columns[0].Num(0) == 1.5 && read.Load().Columns[1].Num(0) == 45000 && read.Load().Columns[2].Text(0) == "مرحبا" && read.Load().Columns[3].Num(0) == 1 && read.Load().Columns.All(c => c.IsNull(1)) && read.Meta["test"] == "yes", "");
        TableFile.Write(root, "empty.xvt", new Table("empty", Array.Empty<Column>()), new Dictionary<string,string>());
        check("empty table", TableFile.Open(root, "empty.xvt").Load().Rows == 0, "");
        TableFile.Write(root, "one.xvt", new Table("one", new Column[] { TextColumn.From("x", new string?[] {"one"}) }), new Dictionary<string,string>());
        check("one row", TableFile.Open(root, "one.xvt").Load().Columns[0].Text(0) == "one", "");
        using (var part = root.OpenWrite("unfinished.xvt")) part.WriteByte(42);
        bool missing = false; try { TableFile.Open(root, "unfinished.xvt"); } catch (FileNotFoundException) { missing = true; }
        check("uncommitted part ignored", missing, "");
        using (var stream = root.OpenWrite("oversize.xvt")) using (var writer = new BinaryWriter(stream)) {
            var header = System.Text.Encoding.UTF8.GetBytes("{\"name\":\"large\",\"rows\":3000000000,\"columns\":[],\"meta\":{}}");
            writer.Write(System.Text.Encoding.ASCII.GetBytes("XVTABLE1")); writer.Write(1); writer.Write(header.Length); writer.Write(header);
        }
        root.Commit("oversize.xvt");
        bool refused = false; try { TableFile.Open(root, "oversize.xvt"); } catch (ExcelApp.UserError) { refused = true; }
        check("oversized table refused", refused, "");
        var values = new double[10_000_000];
        var large = new Table("speed", Enumerable.Range(0, 10).Select(i => (Column)new NumberColumn("c" + i, values, null)).ToArray());
        var timer = System.Diagnostics.Stopwatch.StartNew();
        TableFile.Write(root, "speed.xvt", large, new Dictionary<string,string>()); timer.Stop();
        double seconds = timer.Elapsed.TotalSeconds; timer.Restart(); TableFile.Open(root, "speed.xvt"); timer.Stop();
        double ms = timer.Elapsed.TotalMilliseconds;
        Console.WriteLine($"SPEED xvt write 10Mx10 {seconds:F3} s, open header {ms:F3} ms");
        check("xvt performance", seconds < 10 && ms < 50, ""); root.Delete("speed.xvt");
        using (var queue = new WorkQueue()) {
            var order = new System.Collections.Concurrent.ConcurrentQueue<int>();
            var done = new TaskCompletionSource<bool>();
            queue.Enqueue("first", async _ => { await Task.Delay(30); order.Enqueue(1); });
            queue.Enqueue("second", _ => { order.Enqueue(2); done.SetResult(true); return Task.CompletedTask; });
            bool finished = done.Task.Wait(TimeSpan.FromSeconds(5));
            check("queue serial order", finished && order.SequenceEqual(new[] {1,2}), "");
        }
        var versions = new Versions(root, "synthetic-" + Guid.NewGuid().ToString("N"));
        string period="2026-10"; var time=new DateTime(2026,10,1);
        versions.Add(period,new VersionInfo { LoadedAt=time, VersionId=1 }); versions.Lock(period,true);
        versions.Add(period,new VersionInfo { LoadedAt=time.AddDays(1), VersionId=2 }); versions.Add(period,new VersionInfo { LoadedAt=time.AddDays(2), VersionId=3 });
        var periodState=versions.Snapshot().Periods[period];
        check("version latest",Versions.Select(periodState,"latest",null,null)?.Version==3,"");
        check("version locked",Versions.Select(periodState,"locked",null,null)?.Version==1 && periodState.Versions.Last().PostClose,"");
        check("version as of",Versions.Select(periodState,null,time.AddHours(25),null)?.Version==2,"");
        versions.Rollback(period,3); check("version rollback",versions.Snapshot().Periods[period].Current==2,"");
        var providerCatalog = new Catalog(root); string report="report-"+Guid.NewGuid().ToString("N");
        providerCatalog.Update(c => { c.ReportTypes.Add(report); return 0; });
        var provider = new TableProvider(root,providerCatalog);
        Table Make(double v) => new("synthetic",new Column[] {new NumberColumn("value",new[]{v},null)});
        provider.SaveAsVersion(report,"2026-10",Make(1),"");
        new Versions(root,report).Lock("2026-10",true);
        provider.SaveAsVersion(report,"2026-10",Make(2),"");
        provider.SaveAsVersion(report,"2026-11",Make(3),"");
        check("provider latest union",provider.Get(new TableRef("report",report))["value"]!.Num(0)==2 && provider.Get(new TableRef("report",report)).Rows==2,"");
        check("provider locked union",provider.Get(new TableRef("report",report,Truth:"locked"))["value"]!.Num(0)==1,"");
        check("provider selected version",provider.Get(new TableRef("report",report,Version:1))["value"]!.Num(0)==1,"");
        string workspaceId="w"+Guid.NewGuid().ToString("N"); string relative="tables/synthetic/"+workspaceId+".xvt";
        TableFile.Write(root,relative,Make(9),new Dictionary<string,string>());
        var workspaceSource=new SourceInfo(987654,"table","synthetic","",0,"synthetic.xlsx","Sheet",1,DateTime.Now,"",new Dictionary<string,int>());
        providerCatalog.Update(c=>{c.Workspace.Add(new WorkspaceEntry(new TableInfo("table",workspaceId,"synthetic","project",1,Array.Empty<ColumnInfo>(),DateTime.Now),relative,workspaceSource));return 0;});
        check("workspace lookup and source",provider.Get(new TableRef("table",workspaceId))["value"]!.Num(0)==9 && provider.SourceOf(987654)?.GroupId=="synthetic","");
        string unionReport="union-"+Guid.NewGuid().ToString("N");providerCatalog.Update(c=>{c.ReportTypes.Add(unionReport);return 0;});
        var million=new double[1_000_000];var millionTable=new Table("synthetic",new Column[]{new NumberColumn("value",million,null)});
        for(int month=1;month<=12;month++) provider.SaveAsVersion(unionReport,$"2025-{month:D2}",millionTable,"");
        var unionTimer=System.Diagnostics.Stopwatch.StartNew();var union=provider.Get(new TableRef("report",unionReport));unionTimer.Stop();
        Console.WriteLine($"SPEED union 12x1M {unionTimer.Elapsed.TotalSeconds:F3} s");check("union performance",union.Rows==12_000_000 && unionTimer.Elapsed.TotalSeconds<3,"");
        for(int month=1;month<=12;month++) root.Delete($"reports/{unionReport}/2025-{month:D2}/v1.xvt"); root.Delete($"reports/{unionReport}/versions.json"); providerCatalog.Update(c=>{c.ReportTypes.Remove(unionReport);return 0;});
        using(var exportQueue=new WorkQueue()) {var exporter=new Exporter(root,exportQueue,()=>new Settings());bool pdfDenied=false;try {exporter.ExportAsync(new ExportJob {Pdf=true},null,CancellationToken.None);} catch(ExcelApp.UserError) {pdfDenied=true;}check("export pdf disabled",pdfDenied,"");}
        var catalog = new Catalog(root); int a = catalog.NextVersion(), b = catalog.NextVersion(); check("catalog version sequence", b == a + 1, "");
    }
    static void Review(Action<string, bool, string> check)
    {
        var root = StoreRoot.ForTest(@"D:\ExcelAppTest\syn-review-store-" + Guid.NewGuid().ToString("N"));
        var ids = new System.Collections.Concurrent.ConcurrentBag<int>();
        Parallel.For(0, 20, _ => ids.Add(new Catalog(root).NextVersion()));
        check("review catalog concurrent instances", ids.Distinct().Count() == 20, "");
        check("review locks require no direct lock files", root.List("", "*.lock").Count == 0, "");
        root.WriteAllText("broken.xvt", "not a table");
        bool plain = false;
        try { TableFile.Open(root, "broken.xvt"); } catch (ExcelApp.UserError) { plain = true; } catch { }
        check("review broken xvt has plain error", plain, "");
        // A header that opens but promises more rows than its column holds fails on the lazy column read: still a plain sentence.
        TableFile.Write(root, "short.xvt", new Table("short", new Column[] { new NumberColumn("N", new[] { 1d, 2d }, null) }), new Dictionary<string, string>());
        byte[] bytes; using (var rs = root.OpenRead("short.xvt")) { bytes = new byte[rs.Length]; rs.ReadExactly(bytes); }
        string text = System.Text.Encoding.Latin1.GetString(bytes);
        int at = text.IndexOf("\"rows\":2", StringComparison.Ordinal); bytes[at + 7] = (byte)'3';
        using (var ws = root.OpenWrite("short.xvt")) ws.Write(bytes); root.Commit("short.xvt");
        bool lazyPlain = false;
        try { TableFile.Open(root, "short.xvt").Load(); } catch (ExcelApp.UserError) { lazyPlain = true; } catch { }
        check("review damaged column has plain error", lazyPlain, "");
        var union = TableProvider.Union("synthetic", new[] {
            new Table("old", new Column[] { new NumberColumn("Code", new[] { 12d }, null) }),
            new Table("new", new Column[] { TextColumn.From("Code", new[] { "0012" }) }) });
        check("review union preserves text schema and leading zeros", union["Code"]!.Type == ColType.Text && union["Code"]!.Text(1) == "0012", "");
    }

}

