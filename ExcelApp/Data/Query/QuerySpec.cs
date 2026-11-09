using System.Text.Json;

namespace ExcelApp.Data.Query;

// Predicate objects retain their wire representation, including their op discriminator.
public sealed class QuerySpec
{
    public TableRef Source { get; set; } = new("table", "");
    public List<JoinSpec>? Joins { get; set; }
    public JsonElement? Where { get; set; }
    public List<Computed>? Computed { get; set; }
    public List<string>? GroupBy { get; set; }
    public List<Aggregate>? Aggregates { get; set; }
    public JsonElement? Having { get; set; }
    public List<SortSpec>? Sort { get; set; }
    public int? Limit { get; set; }
    public List<string>? Columns { get; set; }
}
public sealed record Computed(string Name, string Expr);
public sealed record Aggregate(string Fn, string? Col, string As);
public sealed record SortSpec(string Col, bool Desc);
public sealed record JoinKey(string Left, string Right);
public sealed class JoinSpec
{
    public TableRef Right { get; set; } = new("table", "");
    public List<JoinKey> On { get; set; } = new();
    public string Kind { get; set; } = "inner";
    public string? Fuzzy { get; set; }
    public double? MinScore { get; set; }
    public string? Prefix { get; set; }
}

public static class RelativeDate
{
    public static bool TryResolve(string value, DateTime now, out DateTime result)
    {
        var today = now.Date;
        var week = today.AddDays(-((int)today.DayOfWeek + 6) % 7);
        var month = new DateTime(today.Year, today.Month, 1);
        result = value switch
        {
            "today" => today, "startOfWeek" => week, "startOfLastWeek" => week.AddDays(-7),
            "startOfMonth" => month, "startOfLastMonth" => month.AddMonths(-1),
            "endOfLastMonth" => month.AddDays(-1), "startOfYear" => new DateTime(today.Year, 1, 1),
            _ => default
        };
        if (result != default) return true;
        if (value.StartsWith("today-", StringComparison.Ordinal) && int.TryParse(value.AsSpan(6), out int days) && days >= 0 && days <= (today - DateTime.MinValue).Days)
        { result = today.AddDays(-days); return true; }
        return false;
    }
}
