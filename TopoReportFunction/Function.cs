using Amazon.Lambda.APIGatewayEvents;
using Amazon.Lambda.Core;
using Newtonsoft.Json;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using System.Globalization;
using Topo.Model.Program;
using Topo.Model.Members;
using Topo.Model.Milestone;
using Topo.Model.OAS;
using Topo.Model.ReportGeneration;
using Topo.Model.SIA;
using Topo.Model.Logbook;
using Topo.Model.Wallchart;
using Topo.Model.AdditionalAwards;
using Topo.Services;
using Topo.Model.Approvals;
using Topo.Model.Progress;
using System;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace TopoReportFunction;

public class Function
{
    private ReportService reportService = new ReportService();

    private JsonSerializerSettings _settings = new JsonSerializerSettings
    {
        DateParseHandling = DateParseHandling.None
    };

    /// <summary>
    /// Returns a Member List Report in a base 64 string of the PDF or XLSX
    /// </summary>
    /// <param name="reportData">A ReportGenerationRequest object containing the report request</param>
    /// <param name="context"></param>
    /// <returns></returns>
    public string FunctionHandler(APIGatewayProxyRequest request, ILambdaContext context)
    {
        try
        {
            CultureInfo cultureInfo = new("en-AU");
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
            CultureInfo.DefaultThreadCurrentUICulture = cultureInfo;
            string? syncfusionLicense = Environment.GetEnvironmentVariable("SyncfusionLicense");
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(syncfusionLicense ?? "");

            var requestBody = request.Body;
            Console.WriteLine(requestBody);
            var reportGenerationRequest = JsonConvert.DeserializeObject<ReportGenerationRequest>(requestBody, _settings);
            if (reportGenerationRequest != null)
            {
                var workbook = reportService.CreateWorkbookWithSheets(1);
                reportService.CurrentUtcOffset = reportGenerationRequest.CurrentUtcOffset;
                switch (reportGenerationRequest.ReportType)
                {
                    case ReportType.MemberList:
                        workbook = GenerateMemberListWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.PatrolList:
                        workbook = GeneratePatrolListWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.PatrolSheets:
                        workbook = GeneratePatrolSheetsWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.SignInSheet:
                        workbook = GenerateSignInSheetWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.EventAttendance:
                        workbook = GenerateEventAttendanceWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.Attendance:
                        workbook = GenerateAttendanceReportWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.OASWorksheet:
                        workbook = GenerateOASWorksheetWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.SIA:
                        workbook = GenerateSIAWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.Milestone:
                        workbook = GenerateMilestoneWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.Logbook:
                        workbook = GenerateLogbookWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.Wallchart:
                        workbook = GenerateWallchartWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.AdditionalAwards:
                        workbook = GenerateAdditionalAwardsWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.Approvals:
                        workbook = GenerateApprovalsWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.PersonalProgress:
                        workbook = GenerateProgressWorkbook(reportGenerationRequest);
                        break;
                    case ReportType.TermProgram:
                        workbook = GenerateTermProgramWorkbook(reportGenerationRequest);
                        break;
                }

                Console.WriteLine("Workbook completed");

                if (workbook != null)
                {
                    MemoryStream strm = new MemoryStream();
                    workbook.Version = ExcelVersion.Excel2016;

                    if (reportGenerationRequest.OutputType == OutputType.PDF)
                    {
                        //Initialize XlsIO renderer.
                        XlsIORenderer renderer = new XlsIORenderer();

                        //Convert Excel document into PDF document 
                        PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);
                        pdfDocument.Save(strm);

                        Console.WriteLine("Workbook streamed to PDF");
                    }

                    if (reportGenerationRequest.OutputType == OutputType.Excel)
                    {
                        //Stream as Excel file
                        workbook.SaveAs(strm);

                        Console.WriteLine("Workbook streamed to Excel");
                    }

                    // return stream in browser
                    return Convert.ToBase64String(strm.ToArray());

                }
            }
        }

        catch(Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
        

        return "";
    }

    private IWorkbook GenerateMemberListWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<MemberListModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateMemberListWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GeneratePatrolListWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<MemberListModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GeneratePatrolListWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.IncludeLeaders);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GeneratePatrolSheetsWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<MemberListModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GeneratePatrolSheetsWorkbook(reportData, reportGenerationRequest.Section);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateSignInSheetWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<MemberListModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateSignInSheetWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.EventName);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateEventAttendanceWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<EventListModel>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateEventAttendanceWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateAttendanceReportWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<AttendanceReportModel>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateAttendanceReportWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.FromDate, reportGenerationRequest.ToDate
                , reportGenerationRequest.OutputType == OutputType.PDF);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateOASWorksheetWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<OASWorksheetAnswers>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateOASWorksheetWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.OutputType == OutputType.PDF, reportGenerationRequest.FormatLikeTerrain
                , reportGenerationRequest.BreakByPatrol);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateSIAWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<SIAProjectListModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateSIAWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.OutputType == OutputType.PDF);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateMilestoneWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<MilestoneSummaryListModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateMilestoneWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.OutputType == OutputType.PDF);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateLogbookWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<MemberLogbookReportViewModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateLogbookWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.OutputType == OutputType.PDF);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateWallchartWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<WallchartItemModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateWallchartWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.OutputType == OutputType.PDF, reportGenerationRequest.BreakByPatrol);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateAdditionalAwardsWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<AdditionalAwardsReportDataModel>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateAdditionalAwardsWorkbook(reportData.AwardSpecificationsList, reportData.SortedAdditionalAwardsList, reportData.DistinctAwards, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateProgressWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<ProgressDetailsPageViewModel>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateProgressWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section, reportGenerationRequest.UnitName);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateApprovalsWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<ApprovalsListModel>>(reportGenerationRequest.ReportData);
        if (reportData != null)
        {
            var workbook = reportService.GenerateApprovalsWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName, reportGenerationRequest.FromDate, reportGenerationRequest.ToDate, reportGenerationRequest.GroupByMember
                , reportGenerationRequest.OutputType == OutputType.PDF);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }

    private IWorkbook GenerateTermProgramWorkbook(ReportGenerationRequest reportGenerationRequest)
    {
        var reportData = JsonConvert.DeserializeObject<List<EventListModel>>(reportGenerationRequest.ReportData, _settings);
        if (reportData != null)
        {
            var workbook = reportService.GenerateTermProgramWorkbook(reportData, reportGenerationRequest.GroupName, reportGenerationRequest.Section
                , reportGenerationRequest.UnitName
                , reportGenerationRequest.OutputType == OutputType.Excel);
            return workbook;
        }
        return reportService.CreateWorkbookWithSheets(1);
    }
}
