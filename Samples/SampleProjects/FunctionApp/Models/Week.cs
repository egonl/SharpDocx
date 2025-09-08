using System.Globalization;

namespace FunctionApp.Models;

public class Week(DateTime monday)
{
    public DateTime Monday { get; set; } = monday;

    public int Number => GetIso8601WeekOfYear(Monday);

    public int Year => Number == 1 && Monday.Month == 12 ? Monday.Year + 1 : Monday.Year;

    public string GetDayText(int day)
    {
        return $"{Monday.AddDays(day):ddd d-M}";
    }

    public static int GetIso8601WeekOfYear(DateTime time)
    {
        var day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
        if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday) time = time.AddDays(3);

        return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(
            time,
            CalendarWeekRule.FirstFourDayWeek,
            DayOfWeek.Monday);
    }
}