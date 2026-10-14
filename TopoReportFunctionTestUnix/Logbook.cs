using Newtonsoft.Json;
using Topo.Model.ReportGeneration;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class Logbook
    {
        private ReportGenerationRequest _reportGenerationRequest;
        private Function _function;

        [TestInitialize]
        public void SetUp()
        {
            SetEnvironment.SetSyncfusionLicense();
            _function = new Function();
            string reportData = File.ReadAllText("ReportData/Logbook.json");
            _reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.Logbook,
                GroupName = "Epping Scout Group",
                Section = "scout",
                UnitName = "Scout Unit",
                OutputType = OutputType.PDF,
                ReportData = reportData
            };
            if (!Directory.Exists("TestResults"))
                Directory.CreateDirectory("TestResults");
        }

        [TestMethod]
        public void Logbook_ToExcel()
        {
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            _reportGenerationRequest.OutputType = OutputType.Excel;
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/Logbook_ToExcel.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

    }
}
