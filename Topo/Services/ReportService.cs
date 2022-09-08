using Newtonsoft.Json;
using System.Text;
using System.Text.Json;
using Topo.Model.ReportGeneration;

namespace Topo.Services
{
    public interface IReportService
    {
        public Task<byte[]> GetMemberListReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList);
        public Task<byte[]> GetPatrolListReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList, bool includeLeaders);
        public Task<byte[]> GetPatrolSheetsReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList);
        public Task<byte[]> GetSignInSheetReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList, string eventName);
        public Task<byte[]> GetEventAttendanceReport(string groupName, string section, string unitName, OutputType outputType, string serialisedEventListModel);
        public Task<byte[]> GetAttendanceReport(string groupName, string section, string unitName, OutputType outputType, string serialisedAttendanceReportData, DateTime fromDate, DateTime toDate);
        public Task<byte[]> GetOASWorksheetReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberAnswers, bool breakByPatrol);
        public Task<byte[]> GetSIAReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
        public Task<byte[]> GetMilestoneReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
        public Task<byte[]> GetLogbookReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
        public Task<byte[]> GetWallchartReport(string groupName, string section, string unitName, OutputType outputType, string serialisedWallchartItems);
    }

    public class ReportService : IReportService
    {
        private readonly HttpClient _httpClient;

        public ReportService(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }
        public async Task<byte[]> GetMemberListReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.MemberList,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedSortedMemberList
            };
            
            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetPatrolListReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList, bool includeLeaders)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.PatrolList,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                IncludeLeaders = includeLeaders,
                OutputType = outputType,
                ReportData = serialisedSortedMemberList
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetPatrolSheetsReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.PatrolSheets,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedSortedMemberList
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetSignInSheetReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList, string eventName)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.SignInSheet,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedSortedMemberList,
                EventName = eventName
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetEventAttendanceReport(string groupName, string section, string unitName, OutputType outputType, string serialisedEventListModel)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.EventAttendance,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedEventListModel
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetAttendanceReport(string groupName, string section, string unitName, OutputType outputType, string serialisedAttendanceReportData, DateTime fromDate, DateTime toDate)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.Attendance,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedAttendanceReportData,
                FromDate = fromDate,
                ToDate = toDate
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetOASWorksheetReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberAnswers, bool breakByPatrol)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.OASWorksheet,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedSortedMemberAnswers,
                BreakByPatrol = breakByPatrol
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetSIAReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.SIA,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedReportData
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetMilestoneReport(string groupName, string section, string unitName, OutputType outputType, string serialisedMilestoneSummaries)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.Milestone,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedMilestoneSummaries
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetWallchartReport(string groupName, string section, string unitName, OutputType outputType, string serialisedWallchartItems)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.Wallchart,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedWallchartItems
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetLogbookReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.Logbook,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedReportData
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        private async Task<byte[]> CallReportGeneratorFunction(ReportGenerationRequest reportGenerationRequest)
        {
            HttpRequestMessage httpRequest = new HttpRequestMessage(HttpMethod.Put, "https://gkgfntjuhcuc3qtjmohaaxigwe0uevae.lambda-url.ap-southeast-2.on.aws/");
            var content = JsonConvert.SerializeObject(reportGenerationRequest);
            httpRequest.Content = new StringContent(content, Encoding.UTF8, "application/json");
            httpRequest.Headers.Add("accept", "application/json, text/plain, */*");
            var response = await _httpClient.SendAsync(httpRequest);
            var responseContent = response.Content.ReadAsStringAsync();
            var result = responseContent.Result;

            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(result);
            return bytes;
        }
    }
}
