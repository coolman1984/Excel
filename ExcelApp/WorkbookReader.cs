// Owns a new hidden Excel instance. Passwords and read-only prompts must never block a build.
using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelViewer.Core;

namespace ExcelApp;

public static class WorkbookReader
{
    public static BuildResult Build(string source, Action<string> report, Func<bool> cancelled, double? stopAt = null)
    {
        if (Machine.Commit() >= Math.Min(stopAt ?? Machine.StopAt, Machine.StopAt)) return new BuildResult { Code = 90, Stopped = "memory" };
        dynamic? xl = null, wb = null;
        var result = new BuildResult { Code = 1 }; var timer = Stopwatch.StartNew();
        try
        {
            if (cancelled()) return new BuildResult { Code = 2, Stopped = "cancelled" };
            report("OPENING");
            xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
            xl.Visible = false; xl.DisplayAlerts = false; xl.AutomationSecurity = 3;
            xl.ScreenUpdating = false; xl.EnableEvents = false; xl.Interactive = false;
            double started = timer.Elapsed.TotalSeconds; timer.Restart();
            wb = xl.Workbooks.Open(source, UpdateLinks: 0, ReadOnly: true, Notify: false, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true);
            double opened = timer.Elapsed.TotalSeconds; report("OPENED");
            result = CacheBuilder.BuildFromOpenWorkbook((object)xl, (object)wb, source, CachePolicy.PathFor(source), report, cancelled, StyleStore.AfterRead);
            result.ExcelStartSec = started; result.OpenSec = opened;
        }
        catch (Exception e) { result.Code = 1; result.Stopped = e.GetType().Name; }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch (Exception e) { result.Code = 1; result.Stopped = e.GetType().Name; }
            try { if (xl != null) xl.Quit(); } catch (Exception e) { result.Code = 1; result.Stopped = e.GetType().Name; }
            if (wb != null && Marshal.IsComObject(wb)) Marshal.FinalReleaseComObject(wb);
            if (xl != null && Marshal.IsComObject(xl)) Marshal.FinalReleaseComObject(xl);
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        }
        return result;
    }
}
