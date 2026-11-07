// Timings and counts only. Never a cell value, a sheet name or a real file name.
using System.Globalization;

namespace ExcelApp;

public static class Log
{
    static readonly object Gate = new();
    static string? folder;
    public static string Folder
    {
        get => folder ??= Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelApp");
        set => folder = value;
    }
    public static string FilePath => Path.Combine(Folder, "log.txt");
    const long MaxBytes = 1 << 20;

    public static void Write(string line)
    {
        try
        {
            lock (Gate)
            {
                Directory.CreateDirectory(Folder);
                var f = new FileInfo(FilePath);
                if (f.Exists && f.Length > MaxBytes)
                {
                    string old = FilePath + ".1";
                    if (File.Exists(old)) File.Delete(old);
                    File.Move(FilePath, old);
                }
                File.AppendAllText(FilePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) + " " + line + Environment.NewLine);
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }
}
