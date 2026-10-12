using Newtonsoft.Json;
using Topo.Model.ReportGeneration;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class Progress
    {
        private ReportGenerationRequest _reportGenerationRequest;
        private Function _function;

        [TestInitialize]
        public void SetUp()
        {
            SetEnvironment.SetSyncfusionLicense();
            _function = new Function();
            string reportData = File.ReadAllText("ReportData/Progress.json");
            _reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.PersonalProgress,
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
        public void Progress_ToPDF()
        {
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/Progress_ToPDF.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void Progress_ToExcel()
        {
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            _reportGenerationRequest.OutputType = OutputType.Excel;
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/Progress_ToExcel.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

    }
}
