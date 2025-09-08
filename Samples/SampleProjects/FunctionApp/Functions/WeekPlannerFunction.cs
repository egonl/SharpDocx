using System.Net;
using FunctionApp.Extensions;
using FunctionApp.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using SharpDocx;

namespace FunctionApp.Functions;

public class WeekPlannerFunction(ILogger<WeekPlannerFunction> logger)
{
    [Function("WeekPlannerFunction")]
    [OpenApiOperation(
        nameof(WeekPlannerFunction),
        ["Calendars"],
        Summary = "Week Planner",
        Description = "Create a calendar for week planning")]
    [OpenApiParameter(
        "year",
        In = ParameterLocation.Query,
        Description = "Year to plan",
        Type = typeof(int),
        Required = false)]
    [OpenApiResponseWithBody(
        HttpStatusCode.OK,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        typeof(byte[]),
        Description = "Returns the calendar as a Word document (.docx)")]
    public IActionResult Run([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequest req)
    {
        logger.LogInformation("Creating week planning");

        var year = req.Query["year"].ToString().ConvertToIntOrDefault(2026);
        var planningSheet = new PlanningSheet($"Planning {year}", year);
        var functionAppDirectory = AppContext.BaseDirectory;
        var viewPath = Path.Combine(functionAppDirectory, "Views/PlanningSheet.cs.docx");
        var document = DocumentFactory.Create(viewPath, planningSheet);
        var stream = document.Generate();

        return new FileStreamResult(stream,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        {
            FileDownloadName = $"Planning {year}.docx"
        };
    }
}
