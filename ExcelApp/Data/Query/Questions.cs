using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace ExcelApp.Data.Query;

public sealed record Question(string Text, QuerySpec? Spec, string? Pattern, string Explanation, bool Ok);
public sealed class Questions
{
    readonly DataServices services;
    public Questions(DataServices services) => this.services = services;
    static string Normalize(string s) => Regex.Replace(s.ToLowerInvariant(), @"[^a-z0-9]+", " ").Trim();
    public Question Ask(string text)
    {
        var tables = services.Tables.List(); if (string.IsNullOrWhiteSpace(text) || tables.Count == 0) return new(text, null, null, "No table could be identified.", false);
        string q = Normalize(text); TableInfo? MatchTable(string value) => tables.Where(t => Normalize(t.Name).Length > 0 && (Normalize(value).Contains(Normalize(t.Name)) || Normalize(t.Name).Contains(Normalize(value)))).OrderByDescending(t => Normalize(t.Name).Length).FirstOrDefault();
        var table = tables.Where(t => q.Contains(Normalize(t.Name)) && Normalize(t.Name).Length > 0).OrderByDescending(t => Normalize(t.Name).Length).FirstOrDefault() ?? tables[0];
        TableRef Ref(TableInfo t) => new(t.Kind, t.Id); var spec = new QuerySpec { Source = Ref(table) }; var columnNames = table.Columns.Where(c => !c.Name.StartsWith('_')).Select(c => c.Name).ToArray();
        string? Col(string value) { string norm = Normalize(value); return columnNames.Where(c => Normalize(c) == norm || (norm.Length > 0 && (norm.Contains(Normalize(c)) || Normalize(c).Contains(norm)))).OrderByDescending(c => Normalize(c) == norm).ThenByDescending(c => Normalize(c).Length).FirstOrDefault(); }
        Question Good(string pattern)
        {
            var date = table.Columns.FirstOrDefault(c => c.Type == ColType.Date)?.Name; var now = DateTime.Today; DateTime? start = null, end = null;
            if (q.Contains("last month")) { end = new DateTime(now.Year, now.Month, 1); start = end.Value.AddMonths(-1); }
            else if (q.Contains("this month")) { start = new DateTime(now.Year, now.Month, 1); end = start.Value.AddMonths(1); }
            else if (q.Contains("last week")) { end = now.AddDays(-((int)now.DayOfWeek + 6) % 7); start = end.Value.AddDays(-7); }
            else if (q.Contains("this week")) { start = now.AddDays(-((int)now.DayOfWeek + 6) % 7); end = start.Value.AddDays(7); }
            else if (q.Contains("yesterday")) { start = now.AddDays(-1); end = now; }
            else if (q.Contains("today")) { start = now; end = now.AddDays(1); }
            else if (q.Contains("this year")) { start = new DateTime(now.Year, 1, 1); end = start.Value.AddYears(1); }
            var days = Regex.Match(q, @"last (\d+) days"); if (days.Success && int.TryParse(days.Groups[1].Value, out int n) && n > 0 && n < 100000) { start = now.AddDays(1 - n); end = now.AddDays(1); }
            if (pattern != "month comparison" && date != null && start.HasValue && end.HasValue)
            {
                var conditions = new List<object>(); if (spec.Where is { } existing) conditions.Add(existing); conditions.Add(new { op = "ge", col = date, value = start.Value.ToOADate() }); conditions.Add(new { op = "lt", col = date, value = end.Value.ToOADate() }); spec.Where = JsonSerializer.SerializeToElement(new { op = "and", items = conditions }, Json.Options);
            }
            return new(text, spec, pattern, "I understood " + pattern + ".", true);
        }
        Question Bad(string why) => new(text, null, null, why, false);
        static JsonElement Pred(object p) => JsonSerializer.SerializeToElement(p, Json.Options);
        void Aggregate(string fn, string? col, string? group) { if (group != null) spec.GroupBy = new() { group }; spec.Aggregates = new() { new(fn, col, "value") }; }
        static object Literal(string value) => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double n) ? n : value.Trim(' ', '\'', '"');
        Match m;
        if ((m = Regex.Match(text, @"^compare\s+(.+?)\s+and\s+(.+?)\s+by\s+(.+)$", RegexOptions.IgnoreCase)).Success)
        {
            var left = MatchTable(m.Groups[1].Value); var right = MatchTable(m.Groups[2].Value); if (left == null || right == null) return Bad("I could not identify both tables.");
            var key = left.Columns.FirstOrDefault(c => Normalize(c.Name) == Normalize(m.Groups[3].Value))?.Name; var rightKey = right.Columns.FirstOrDefault(c => Normalize(c.Name) == Normalize(m.Groups[3].Value))?.Name;
            if (key == null || rightKey == null) return Bad("I could not identify a shared key."); spec.Source = Ref(left); spec.Joins = new() { new JoinSpec { Right = Ref(right), Kind = "left", On = new() { new(key, rightKey) } } }; return Good("compare tables");
        }
        if ((m = Regex.Match(text, @"^what changed in\s+([A-Za-z]+)(?:\s+(\d{4}))?$", RegexOptions.IgnoreCase)).Success)
        {
            string monthName = m.Groups[1].Value; int month = Enumerable.Range(1, 12).FirstOrDefault(n => CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(n).Equals(monthName, StringComparison.OrdinalIgnoreCase) || CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName(n).Equals(monthName, StringComparison.OrdinalIgnoreCase));
            var report = table.Kind == "report" ? table : tables.FirstOrDefault(t => t.Kind == "report"); if (month == 0 || report == null) return Bad("I could not identify the report and month.");
            int year = m.Groups[2].Success ? int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) : DateTime.Today.Year - (month > DateTime.Today.Month ? 1 : 0); spec.Source = Ref(report); spec.History = new HistoryRequest(report.Id, $"{year:D4}-{month:D2}"); return Good("history diff");
        }
        if ((m = Regex.Match(text, @"^(.+?)\s+this month vs last month$", RegexOptions.IgnoreCase)).Success)
        {
            var measure = Col(m.Groups[1].Value); var date = table.Columns.FirstOrDefault(c => c.Type == ColType.Date)?.Name; if (measure == null || date == null) return Bad("I could not identify the measure and date column.");
            spec.Computed = new() { new("Month", "month([" + date + "])") }; spec.GroupBy = new() { "Month" }; spec.Aggregates = new() { new("sum", measure, "value") }; spec.Where = Pred(new { op = "and", items = new object[] { new { op = "ge", col = date, value = "startOfLastMonth" }, new { op = "lt", col = date, value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).ToOADate() } } }); return Good("month comparison");
        }
        if ((m = Regex.Match(text, @"^how many rows in\s+(.+)$", RegexOptions.IgnoreCase)).Success) { var t = MatchTable(m.Groups[1].Value); if (t == null) return Bad("I could not identify that table."); spec.Source = Ref(t); Aggregate("count", null, null); return Good("count rows"); }
        if ((m = Regex.Match(text, @"^show\s+(.+)$", RegexOptions.IgnoreCase)).Success) { var t = MatchTable(m.Groups[1].Value); if (t == null) return Bad("I could not identify that table."); spec.Source = Ref(t); return Good("show table"); }
        if ((m = Regex.Match(text, @"^last\s+(\d+)\s+days\s+of\s+(.+)$", RegexOptions.IgnoreCase)).Success)
        {
            var t = MatchTable(m.Groups[2].Value); var date = t?.Columns.FirstOrDefault(c => c.Type == ColType.Date)?.Name; if (t == null || date == null) return Bad("I could not identify a table and date column."); spec.Source = Ref(t); int days = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture); spec.Where = Pred(new { op = "ge", col = date, value = "today-" + Math.Max(0, days - 1) }); return Good("last days");
        }
        if ((m = Regex.Match(text, @"^(top|bottom)\s+(\d+)\s+(.+?)\s+by\s+(.+)$", RegexOptions.IgnoreCase)).Success)
        {
            var group = Col(m.Groups[3].Value); var measure = Col(m.Groups[4].Value); if (group == null || measure == null) return Bad("I could not identify the group and measure."); Aggregate("sum", measure, group); spec.Sort = new() { new("value", m.Groups[1].Value.Equals("top", StringComparison.OrdinalIgnoreCase)) }; spec.Limit = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture); return Good(m.Groups[1].Value.ToLowerInvariant());
        }
        if ((m = Regex.Match(text, @"^(total|average|max|min)\s+(.+?)(?:\s+by\s+(.+))?$", RegexOptions.IgnoreCase)).Success)
        {
            var measure = Col(m.Groups[2].Value); var group = m.Groups[3].Success ? Col(m.Groups[3].Value) : null; if (measure == null || (m.Groups[3].Success && group == null)) return Bad("I could not identify the requested columns."); string fn = m.Groups[1].Value.ToLowerInvariant() switch { "total" => "sum", "average" => "avg", var x => x }; Aggregate(fn, measure, group); return Good(fn);
        }
        if ((m = Regex.Match(text, @"^count\s+by\s+(.+)$", RegexOptions.IgnoreCase)).Success) { var group = Col(m.Groups[1].Value); if (group == null) return Bad("I could not identify the group."); Aggregate("count", null, group); return Good("count by group"); }
        if ((m = Regex.Match(text, @"^(?:rows where\s+)?(.+?)\s*(>=|<=|<>|>|<|=|\bis\b)\s*(.+)$", RegexOptions.IgnoreCase)).Success)
        {
            var col = Col(m.Groups[1].Value); if (col == null) return Bad("I could not identify the filter column."); string op = m.Groups[2].Value.ToLowerInvariant() switch { ">" => "gt", "<" => "lt", ">=" => "ge", "<=" => "le", "<>" => "ne", _ => "eq" }; spec.Where = Pred(new { op, col, value = Literal(m.Groups[3].Value) }); return Good("filter rows");
        }
        if ((m = Regex.Match(text, @"^duplicates in\s+(.+)$", RegexOptions.IgnoreCase)).Success) { var col = Col(m.Groups[1].Value); if (col == null) return Bad("I could not identify the column."); spec.Where = Pred(new { op = "duplicate", cols = new[] { col } }); return Good("duplicates"); }
        if ((m = Regex.Match(text, @"^(?:blank|missing)\s+(.+)$", RegexOptions.IgnoreCase)).Success) { var col = Col(m.Groups[1].Value); if (col == null) return Bad("I could not identify the column."); spec.Where = Pred(new { op = "blank", col }); return Good("blank values"); }
        if ((m = Regex.Match(text, @"^sort by\s+(.+?)(?:\s+(desc|asc))?$", RegexOptions.IgnoreCase)).Success) { var col = Col(m.Groups[1].Value); if (col == null) return Bad("I could not identify the sort column."); spec.Sort = new() { new(col, m.Groups[2].Value.Equals("desc", StringComparison.OrdinalIgnoreCase)) }; return Good("sort"); }
        if ((m = Regex.Match(text, @"^(.+?)\s+per\s+(.+?)\s+in\s+(.+)$", RegexOptions.IgnoreCase)).Success) { var measure = Col(m.Groups[1].Value); var group = Col(m.Groups[2].Value); if (measure == null || group == null) return Bad("I could not identify the requested columns."); Aggregate("sum", measure, group); return Good("measure per group"); }
        if ((m = Regex.Match(text, @"^trend of (.+) by month$", RegexOptions.IgnoreCase)).Success) { var measure = Col(m.Groups[1].Value); var date = table.Columns.FirstOrDefault(c => c.Type == ColType.Date)?.Name; if (measure == null || date == null) return Bad("I could not identify the measure and date column."); spec.Computed = new() { new("Month", "month([" + date + "])") }; Aggregate("sum", measure, "Month"); spec.Sort = new() { new("Month", false) }; return Good("monthly trend"); }
        return Bad("I could not match that question to a supported pattern.");
    }
}
