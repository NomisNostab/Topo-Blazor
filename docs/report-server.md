# Topo report server

One local web server that does two jobs, so Topo can run entirely on your own machine:

1. **Serves Topo's pages**, the same Blazor WebAssembly app that is on GitHub Pages.
2. **Stands in for the AWS Lambda function** Topo sends its report requests to, so reports are made locally with no AWS account.

Pages served from it send their report requests back to it, not to AWS. Topo still signs in to Terrain and reads Terrain's data as usual; only the reports are made locally.

It has no report code of its own. It starts the real `TopoReportFunction.Function` and calls `FunctionHandler`, the same method AWS calls, so what it returns is what the deployed function would return. When a page or a report changes, rebuilding (or `dotnet watch`) is all that is needed to see the change here.

- Project: `TopoReportServer/`
- Default address: `http://localhost:5070`
- Needs: the [.NET 10 SDK](https://dotnet.microsoft.com/download). No AWS tools, no Docker.

## Run it

From the root of the repository:

```bash
dotnet run --project TopoReportServer
```

Leave it running, then open **http://localhost:5070/** in a browser. That is Topo: sign in to Terrain as usual and make reports. The server prints a line for every report it makes, and each report you make in Topo from that page is made by the server, not by AWS.

The first run builds Topo as well as the report function, so it takes a little longer. Check the server is up without a browser:

```bash
curl http://localhost:5070/health
```

Make a report from one of the sample requests and open it:

```bash
curl -s -X PUT http://localhost:5070/ \
  -H "Content-Type: application/json" \
  --data @TopoReportServer/samples/member-list-pdf.json | base64 -d > member-list.pdf
```

### Keep it updated while you change code

```bash
dotnet watch --project TopoReportServer
```

`dotnet watch` rebuilds and restarts the server whenever a file in `Topo`, `TopoReportFunction`, `Services` or `Model` changes. Refresh the browser tab to get the new pages (and, if the browser keeps the old copy, hard refresh with Ctrl+Shift+R).

### Other address or port

```bash
dotnet run --project TopoReportServer --urls http://localhost:6000
```

## What the server answers

| Request | Answer |
|---|---|
| `GET /` and any page path (`/members`, `/login`, …) | Topo's pages and the files they load |
| `PUT /` or `POST /` | A report, see below |
| `GET /appsettings.json` | `{"reportFunctionUrl": "http://<this server>/"}`, which is how the pages learn where to send reports |
| `GET /health` | `{"status":"ok","syncfusionLicenceSet":…}` |
| `GET /report-server` | A plain text status page listing the report types this build knows about |

## The report request and the reply

The report call is the same one the Lambda function URL answers:

| | |
|---|---|
| **Method and path** | `PUT /` (`POST /` also works) |
| **Body** | A `ReportGenerationRequest` as JSON |
| **Reply** | `200`, `text/plain`, the finished file as a **base 64 string** |
| **On failure** | `200` with an **empty body**, and the stack trace in the server's console |

The empty reply on failure is deliberate: the real function catches its own errors and returns `""`, so a client sees the same thing here as in AWS.

### Request fields

These come from `Model/ReportGeneration/ReportGenerationRequest.cs`:

| Field | Meaning |
|---|---|
| `ReportType` | Which report. The name of a `ReportType` value (`"MemberList"`) or its number |
| `OutputType` | `"PDF"` or `"Excel"` |
| `GroupName`, `Section`, `UnitName` | Printed in the report heading. `Section` also picks the logo (`Joey`, `Cub`, `Scout`, `Venturer`, `Rover`) |
| `ReportData` | **A string** holding more JSON: the data the report is made from. Its shape depends on `ReportType`, see below |
| `IncludeLeaders`, `EventName`, `FromDate`, `ToDate`, `BreakByPatrol`, `GroupByMember`, `FormatLikeTerrain`, `CurrentUtcOffset` | Options used by some reports. Leave them out when not needed |

`ReportData` is JSON inside a JSON string, so it must be escaped, which is easiest to do by building the request in a script. See `TopoReportServer/samples/member-list-pdf.json` for a working example.

### What `ReportData` holds for each report

The type is what `TopoReportFunction/Function.cs` deserialises for that report. To see the exact fields, open the class under `Model/`.

| `ReportType` | `ReportData` is a JSON… |
|---|---|
| `MemberList`, `PatrolList`, `PatrolSheets`, `SignInSheet` | list of `MemberListModel` |
| `EventAttendance` | `EventListModel` |
| `Attendance` | `AttendanceReportModel` |
| `OASWorksheet` | list of `OASWorksheetAnswers` |
| `SIA` | list of `SIAProjectListModel` |
| `Milestone` | list of `MilestoneSummaryListModel` |
| `Logbook` | list of `MemberLogbookReportViewModel` |
| `Wallchart` | list of `WallchartItemModel` |
| `AdditionalAwards` | `AdditionalAwardsReportDataModel` |
| `PersonalProgress` | `ProgressDetailsPageViewModel` |
| `Approvals` | list of `ApprovalsListModel` |
| `TermProgram` | list of `EventListModel` |

Reports added later follow the same pattern: add a case to the `switch` in `Function.cs` and it is available here with no change to the server. `GET /report-server` lists the report types the running build knows about.

The easiest way to get real data for a report is to copy a request Topo sends: point Topo at the server (below), make the report in Topo, and copy the request body the server prints to its console.

## How the pages find the server

Every report goes through one method, `CallReportGeneratorFunction` in `Topo/Services/ReportService.cs`, and that method gets its address from one place: the `ReportFunctionUrl` property at the top of the class. The property returns the `ReportFunctionUrl` setting when there is one, and otherwise `DefaultReportFunctionUrl`, the AWS address. To change where all reports go, change that setting or that constant, nothing else. A page served by this server needs no code change:

1. Topo has an empty `Topo/wwwroot/appsettings.json`. Blazor only reads settings at start-up if that file exists.
2. This server answers `/appsettings.json` itself, ahead of the static files, with `ReportFunctionUrl` set to its own address.
3. `ReportService` sees the setting and sends reports there.

On GitHub Pages, or anywhere else the empty file is served, there is no setting and Topo uses the AWS addresses exactly as before.

To use the server's report function from a Topo run some other way (for example `dotnet run --project Topo`), put `{ "ReportFunctionUrl": "http://localhost:5070/" }` in your own `Topo/wwwroot/appsettings.json`. Do not commit that change.

The server allows calls from any web page (CORS), which is what lets such a Topo call it from a different address. This is for local development only: do not put the server on a public address. To restrict it, set the allowed addresses in `TopoReportServer/appsettings.json`:

```json
"ReportServer": { "AllowedOrigins": [ "http://localhost:5283" ] }
```

## Syncfusion licence

The reports are made with Syncfusion. The real function reads the licence key from an environment variable called `SyncfusionLicense`, and so does the server:

```bash
SyncfusionLicense="your key" dotnet run --project TopoReportServer
```

Without it, reports are still made but PDFs carry a Syncfusion trial banner across the first page. `GET /health` shows whether a key was found (`syncfusionLicenceSet`). Never commit a key.

## How it differs from AWS

| | AWS Lambda | This server |
|---|---|---|
| Sign-in or key | None on the function URL | None |
| Several requests at once | Several instances | One report at a time; the rest wait |
| Size and time limits | Response size and 5 minute limit | None |
| Fonts and logos | In the function's folder | Copied beside the server from `TopoReportFunction/Fonts` and `Images` |
| Memory, CPU, cold starts | Set in AWS | Your machine's |

Because the Lambda limits do not apply, a report that works here can still be too large or slow in AWS.

## Handing it to another developer

The server lives in the repository, so a clone or a branch that contains `TopoReportServer/` is a complete copy. They need the .NET 10 SDK and the commands above.

To give them something to run without the source, publish it. The result holds the server, the report function and Topo's pages together:

```bash
dotnet publish TopoReportServer -c Release -o report-server-publish
cd report-server-publish && ./TopoReportServer --urls http://localhost:5070     # TopoReportServer.exe on Windows
```

Then open http://localhost:5070/. A published copy is a snapshot: publish again to include later changes. It needs the .NET 10 **runtime** on the machine (or add `--self-contained -r linux-x64`, `win-x64` or `osx-arm64` to the publish to include it).

Avoid ports 5060 and 5061: browsers refuse to load pages from them (they are the SIP ports on Chrome's blocked list).

## Keeping it up to date

Nothing in `TopoReportServer` needs to change when pages or reports do. After pulling changes, `dotnet run` (or `dotnet watch`) rebuilds it. It only needs touching when:

- The function's entry point changes. `Function.FunctionHandler(APIGatewayProxyRequest, ILambdaContext)` is called in `TopoReportServer/Program.cs`.
- The function starts using something from the Lambda context. `TopoReportServer/LocalLambdaContext.cs` supplies placeholder values.
- The function needs more files beside it. Add them to `TopoReportServer/TopoReportServer.csproj` the way `Fonts` and `Images` are.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| Empty reply, `No report came back` in the log | The function threw. Read the stack trace printed just above it, often a `ReportData` that does not match the type for that `ReportType` |
| PDF fails but Excel works | `Fonts` missing next to the server build. Rebuild; check `TopoReportServer/bin/Debug/net10.0/Fonts` exists |
| Topo shows a CORS error | Topo is running at an address not in `AllowedOrigins`, or the server is not running |
| The page loads but reports still come from AWS (no "Report made" line in the server) | The page was not served by this server, or the browser kept an old copy: hard refresh. Check `http://localhost:5070/appsettings.json` shows `reportFunctionUrl` |
| Browser says the page cannot be reached on port 5060 | That port is blocked by browsers (`ERR_UNSAFE_PORT`); use another |
| Pages look unstyled or Syncfusion parts are missing | The server was started from a build without Topo's static files; run it with `dotnet run --project TopoReportServer` or the published copy |
| "Address already in use" | Something else is on port 5070; use `--urls` |
| Trial banner across the PDF | No `SyncfusionLicense` set (see above) |
| A report type is missing from `GET /report-server` | The branch you built does not contain that report yet |
