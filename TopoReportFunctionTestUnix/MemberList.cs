using Newtonsoft.Json;
using Topo.Model.ReportGeneration;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class MemberList
    {
        private ReportGenerationRequest _reportGenerationRequest;
        private Function _function;
        private string _reportDataWithLeaders = "";

        [TestInitialize]
        public void SetUp()
        {
            SetEnvironment.SetSyncfusionLicense();
            _function = new Function();
            string reportData = File.ReadAllText("ReportData/MemberList.json");
           
            _reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.MemberList,
                GroupName = "Epping Scout Group",
                Section = "scout",
                UnitName = "Scout Unit",
                OutputType = OutputType.PDF,
                ReportData = reportData
            };
            _reportDataWithLeaders = File.ReadAllText("ReportData/MemberListWithLeaders.json");
            if (!Directory.Exists("TestResults"))
                Directory.CreateDirectory("TestResults");
        }

        [TestMethod]
        public void MemberList_ToPDF()
        {
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/MemberList_ToPDF.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void MemberList_ToExcel()
        {
            _reportGenerationRequest.OutputType = OutputType.Excel;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/MemberList_ToExcel.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void PatrolList_ToPDF_IncludeLeaders_False()
        {
            _reportGenerationRequest.ReportType = ReportType.PatrolList;
            _reportGenerationRequest.IncludeLeaders = false;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/PatrolList_ToPDF_IncludeLeaders_False.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void PatrolList_ToExcel_IncludeLeaders_False()
        {
            _reportGenerationRequest.ReportType = ReportType.PatrolList;
            _reportGenerationRequest.IncludeLeaders = false;
            _reportGenerationRequest.OutputType = OutputType.Excel;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/PatrolList_ToExcel_IncludeLeaders_False.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void PatrolList_ToPDF_IncludeLeaders_True()
        {
            _reportGenerationRequest.ReportType = ReportType.PatrolList;
            _reportGenerationRequest.IncludeLeaders = true;
            _reportGenerationRequest.ReportData = _reportDataWithLeaders;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/PatrolList_ToPDF_IncludeLeaders_True.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void PatrolList_ToExcel_IncludeLeaders_True()
        {
            _reportGenerationRequest.ReportType = ReportType.PatrolList;
            _reportGenerationRequest.IncludeLeaders = true;
            _reportGenerationRequest.ReportData = _reportDataWithLeaders;
            _reportGenerationRequest.OutputType = OutputType.Excel;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/PatrolList_ToExcel_IncludeLeaders_True.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void PatrolSheets_ToPDF()
        {
            _reportGenerationRequest.ReportType = ReportType.PatrolSheets;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/PatrolSheets_ToPDF.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void PatrolSheets_ToExcel()
        {
            _reportGenerationRequest.ReportType = ReportType.PatrolSheets;
            _reportGenerationRequest.OutputType = OutputType.Excel;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/PatrolSheets_ToPDF.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }
    }
}
