using System.Globalization;
namespace ExcelApp.Data;
public static class Periods
{
    static DateTime StartOfWeek(DateTime date, DayOfWeek first) => date.Date.AddDays(-((7+(int)date.DayOfWeek-(int)first)%7));
    static DateTime FiscalStart(int year, WeekRule rule) {
        if(!DateTime.TryParseExact($"{year}-{rule.FiscalYearStart}","yyyy-MM-dd",CultureInfo.InvariantCulture,DateTimeStyles.None,out var date)) throw new UserError("The fiscal year start is invalid.");
        return StartOfWeek(date,rule.FirstDay);
    }
    public static string PeriodOf(double serial,Granularity granularity,WeekRule rule) {
        var date=DateTime.FromOADate(serial).Date;
        if(granularity==Granularity.Day) return date.ToString("yyyy-MM-dd",CultureInfo.InvariantCulture);
        if(granularity==Granularity.Month) return date.ToString("yyyy-MM",CultureInfo.InvariantCulture);
        if(rule.Numbering != "fiscal") return $"{ISOWeek.GetYear(date):D4}-W{ISOWeek.GetWeekOfYear(date):D2}";
        int year=date.Year; if(date<FiscalStart(year,rule)) year--; else if(date>=FiscalStart(year+1,rule)) year++;
        int week=(int)((date-FiscalStart(year,rule)).TotalDays/7)+1; return $"FY{year:D4}-W{week:D2}";
    }
    public static DateTime LockTime(string period,Granularity granularity,LockRule rule) => LockTime(period,granularity,rule,new WeekRule());
    public static DateTime LockTime(string period,Granularity granularity,LockRule rule,WeekRule week) {
        if(rule.WorkingDay<1) throw new UserError("The working day must be positive.");
        DateTime start;
        if(granularity==Granularity.Day) start=DateTime.ParseExact(period,"yyyy-MM-dd",CultureInfo.InvariantCulture).AddDays(1);
        else if(granularity==Granularity.Month) start=DateTime.ParseExact(period+"-01","yyyy-MM-dd",CultureInfo.InvariantCulture).AddMonths(1);
        else if(period.StartsWith("FY",StringComparison.Ordinal)) { int year=int.Parse(period.AsSpan(2,4),CultureInfo.InvariantCulture); int number=int.Parse(period.AsSpan(8),CultureInfo.InvariantCulture); start=FiscalStart(year,week).AddDays(number*7); }
        else { int year=int.Parse(period.AsSpan(0,4),CultureInfo.InvariantCulture); int number=int.Parse(period.AsSpan(6),CultureInfo.InvariantCulture); start=ISOWeek.ToDateTime(year,number,DayOfWeek.Monday).AddDays(7); }
        var holidays=new HashSet<string>(rule.Holidays,StringComparer.Ordinal); int found=0;
        while(true) { if(start.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday) && !holidays.Contains(start.ToString("yyyy-MM-dd",CultureInfo.InvariantCulture)) && ++found==rule.WorkingDay) return start.AddDays(1).AddSeconds(-1); start=start.AddDays(1); }
    }
}
