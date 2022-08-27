using Newtonsoft.Json;
using System.Text;
using System.Text.Json;
using Topo.Model.ReportGeneration;

namespace Topo.Services
{
    public interface IReportService
    {
        public Task<byte[]> GetMemberListReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList);
        public Task<byte[]> GetPatrolListReport(string groupName, string section, string unitName, bool includeLeaders, OutputType outputType, string serialisedSortedMemberList);
        public Task<byte[]> GetPatrolSheetsReport(string groupName, string section, string unitName, OutputType outputType, string serialisedSortedMemberList);
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

        public async Task<byte[]> GetPatrolListReport(string groupName, string section, string unitName, bool includeLeaders, OutputType outputType, string serialisedSortedMemberList)
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
