using System.Diagnostics;
using System.Text;
using Amazon.Lambda.APIGatewayEvents;
using Topo.Model.ReportGeneration;
using TopoReportServer;

// Runs Topo on this machine: it serves Topo's pages (the Blazor WebAssembly app) and stands in for the AWS Lambda function
// that Topo sends report requests to.
//
//   GET /             Topo's pages (and the rest of the files they need)
//   PUT (or POST) /   body: a ReportGenerationRequest as JSON
//                     reply: the finished PDF or Excel file as a base 64 string (text/plain), like the Lambda's function URL
//   GET /appsettings.json   tells the pages to send report requests back to this server
//   GET /report-server      a plain text status page, GET /health a quick check
//
// It does not copy any report code: it creates the real TopoReportFunction.Function and calls FunctionHandler,
// the same method AWS calls. See docs/report-server.md.

// The reports find their fonts and logos relative to the working directory, as they do on AWS, so run from the build output whatever folder it was started from
Directory.SetCurrentDirectory(AppContext.BaseDirectory);

var builder = WebApplication.CreateBuilder(args);

var allowedOrigins = builder.Configuration.GetSection("ReportServer:AllowedOrigins").Get<string[]>() ?? new[] { "*" };
builder.Services.AddCors(options => options.AddDefaultPolicy(policy =>
{
    // Topo runs in a browser, so its calls here are cross-origin and are preceded by an OPTIONS preflight
    if (allowedOrigins.Contains("*"))
    {
        policy.AllowAnyOrigin();
    }
    else
    {
        policy.WithOrigins(allowedOrigins);
    }
    policy.AllowAnyHeader().AllowAnyMethod();
}));

// Finds the files Topo's pages need that come from packages (Syncfusion styles and scripts), however the server was started
builder.WebHost.UseStaticWebAssets();

var app = builder.Build();
app.UseCors();

// Topo reads this file when it starts (it has an empty one of its own in wwwroot). Answering it here, ahead of the static files,
// is what makes a Topo served from here send its reports back here and not to AWS.
app.Use(async (context, next) =>
{
    if (HttpMethods.IsGet(context.Request.Method) && context.Request.Path == "/appsettings.json")
    {
        await Results.Json(new { ReportFunctionUrl = $"{context.Request.Scheme}://{context.Request.Host}/" }).ExecuteAsync(context);
    }
    else
    {
        await next();
    }
});

app.UseBlazorFrameworkFiles();
app.UseStaticFiles();

// One function instance for the life of the server, as AWS keeps one warm between calls
var function = new TopoReportFunction.Function();
// The function sets culture and licence from process settings, so only one report is made at a time
var oneAtATime = new SemaphoreSlim(1, 1);

app.MapGet("/report-server", () => Results.Text(
    "Topo report server is running.\n\n" +
    "Topo's pages are at /. PUT or POST a ReportGenerationRequest (JSON) to / to get a report back as a base 64 string.\n" +
    "Report types: " + string.Join(", ", Enum.GetNames<ReportType>()) + "\n" +
    "Output types: " + string.Join(", ", Enum.GetNames<OutputType>()) + "\n" +
    "GET /health for a quick check. See docs/report-server.md.\n"));

app.MapGet("/health", () => Results.Ok(new { status = "ok", syncfusionLicenceSet = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("SyncfusionLicense")) }));

app.MapMethods("/", new[] { "PUT", "POST" }, async (HttpRequest http, ILoggerFactory loggers) =>
{
    var log = loggers.CreateLogger("Report");
    using var reader = new StreamReader(http.Body, Encoding.UTF8);
    var body = await reader.ReadToEndAsync();

    var timer = Stopwatch.StartNew();
    string result;
    await oneAtATime.WaitAsync();
    try
    {
        // The Lambda function URL hands the body to the function inside an API Gateway style request
        result = function.FunctionHandler(new APIGatewayProxyRequest { Body = body }, new LocalLambdaContext());
    }
    finally
    {
        oneAtATime.Release();
    }

    // The real function swallows its own errors and returns "", so an empty answer means it failed (its stack trace is in the console)
    if (string.IsNullOrEmpty(result))
    {
        log.LogWarning("No report came back after {Ms} ms (request was {Length} characters). The error, if any, is printed above.", timer.ElapsedMilliseconds, body.Length);
    }
    else
    {
        log.LogInformation("Report made in {Ms} ms: {Chars} base 64 characters from a {Length} character request", timer.ElapsedMilliseconds, result.Length, body.Length);
    }

    return Results.Text(result, "text/plain");
});

// Anything else is a page of Topo (it has its own routes, such as /members), so it gets the app's index.html
app.MapFallbackToFile("index.html");

app.Run();
