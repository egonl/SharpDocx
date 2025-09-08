namespace FunctionApp.Models;

public class PlanningSheet
{
    public string Title;
    public List<Week> Weeks = [];

    public PlanningSheet(string title, int year)
    {
        Title = title;
        AddWeeks(year);
    }

    public void AddWeeks(int year)
    {
        var monday = GetFirstMonday(new DateTime(year, 1, 1));

        var week = new Week(monday);
        if (week.Number == 2)
            // Add Week 1.
            Weeks.Add(new Week(monday.AddDays(-7)));

        while ((week = new Week(monday)).Year == year)
        {
            Weeks.Add(week);
            monday = monday.AddDays(7);
        }
    }

    private static DateTime GetFirstMonday(DateTime startDate)
    {
        var daysUntilMonday = (DayOfWeek.Monday - startDate.DayOfWeek + 7) % 7;
        return startDate.AddDays(daysUntilMonday);
    }
}