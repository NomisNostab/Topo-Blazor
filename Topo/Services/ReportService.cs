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
        public Task<byte[]> GetOASWorksheetReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberAnswers, bool breakByPatrol, bool formatLikeTerrain);
        public Task<byte[]> GetSIAReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
        public Task<byte[]> GetMilestoneReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
        public Task<byte[]> GetLogbookReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
        public Task<byte[]> GetWallchartReport(string groupName, string section, string unitName, OutputType outputType, string serialisedWallchartItems, bool breakByPatrol);
        public Task<byte[]> GetAdditionalAwardsReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
        public Task<byte[]> GetApprovalsReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData, DateTime fromDate, DateTime toDate, bool groupByMember);
        public Task<byte[]> GetProgressReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
        public Task<byte[]> GetTermProgramReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData);
    }

    public class ReportService : IReportService
    {
        private readonly HttpClient _httpClient;
        private readonly ILogger<TerrainAPIService> _logger;
        private readonly StorageService _storageService;

        public ReportService(HttpClient httpClient, ILogger<TerrainAPIService> logger, StorageService storageService)
        {
            _httpClient = httpClient;
            _logger = logger;
            _storageService = storageService;
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

        public async Task<byte[]> GetOASWorksheetReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberAnswers, bool breakByPatrol, bool formatLikeTerrain)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.OASWorksheet,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedSortedMemberAnswers,
                BreakByPatrol = breakByPatrol,
                FormatLikeTerrain = formatLikeTerrain
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

        public async Task<byte[]> GetWallchartReport(string groupName, string section, string unitName, OutputType outputType, string serialisedWallchartItems, bool breakByPatrol = false)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.Wallchart,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedWallchartItems,
                BreakByPatrol = breakByPatrol
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

        public async Task<byte[]> GetAdditionalAwardsReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.AdditionalAwards,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedReportData
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetApprovalsReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData, DateTime fromDate, DateTime toDate, bool groupByMember)
        {
            var reportGenerationRequest = new ReportGenerationRequest() 
            {
                ReportType = ReportType.Approvals,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedReportData,
                FromDate = fromDate,
                ToDate = toDate,
                GroupByMember = groupByMember
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetProgressReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.PersonalProgress,
                GroupName = groupName,
                Section = section,
                UnitName = unitName,
                OutputType = outputType,
                ReportData = serialisedReportData
            };

            var report = await CallReportGeneratorFunction(reportGenerationRequest);
            return report;
        }

        public async Task<byte[]> GetTermProgramReport(string groupName, string section, string unitName, OutputType outputType, string serialisedReportData)
        {
            var reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.TermProgram,
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
#if DEBUG
            string functionUrl = "https://57aqgtwtgggsk47rtvpgqifima0tvhwy.lambda-url.ap-southeast-2.on.aws/";
#else
            string functionUrl = "https://qwhcdbhrempok4kpmk6utzavxq0zjzha.lambda-url.ap-southeast-2.on.aws/";
#endif
            HttpRequestMessage httpRequest = new HttpRequestMessage(HttpMethod.Put, functionUrl);
            var content = JsonConvert.SerializeObject(reportGenerationRequest);
            string message = $"Version: {_storageService.Version}; Date: {DateTime.Now.ToString("dd/MM/yyyy : HH:mm:ss")}; ReportGenerationRequestData: {reportGenerationRequest.ReportData}";
            _logger.LogInformation(message);
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
