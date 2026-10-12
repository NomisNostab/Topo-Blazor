using Newtonsoft.Json;
using Topo.Model.ReportGeneration;
using TopoReportFunction;


namespace TopoReportFunctionTest
{
    [TestClass]
    public class Events
    {
        private ReportGenerationRequest _reportGenerationRequest;
        private Function _function;
        private string _attendanceListReportData = "";
        private string _attendanceReportData = "";
        private string _termProgramReportData = "";

        [TestInitialize]
        public void SetUp()
        {
            SetEnvironment.SetSyncfusionLicense();
            _function = new Function();
            string reportData = File.ReadAllText("ReportData/EventSignIn.json");
            _reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.SignInSheet,
                GroupName = "Epping Scout Group",
                Section = "scout",
                UnitName = "Scout Unit",
                OutputType = OutputType.PDF,
                ReportData = reportData
            };
            _attendanceListReportData = File.ReadAllText("ReportData/EventAttendanceList.json");
            _attendanceReportData = File.ReadAllText("ReportData/EventAttendance.json");
            _termProgramReportData = File.ReadAllText("ReportData/EventTermProgram.json");
            if (!Directory.Exists("TestResults"))
                Directory.CreateDirectory("TestResults");
        }

        [TestMethod]
        public void SignInSheet_ToPDF()
        {
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/SignInSheet_ToPDF.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void AttendanceList_ToExcel()
        {
            _reportGenerationRequest.ReportType = ReportType.EventAttendance;
            _reportGenerationRequest.ReportData = _attendanceListReportData;
            _reportGenerationRequest.OutputType = OutputType.Excel;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/AttendanceList_ToExcel.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void Attendance_ToPDF()
        {
            _reportGenerationRequest.ReportType = ReportType.Attendance;
            _reportGenerationRequest.ReportData = _attendanceReportData;
            _reportGenerationRequest.FromDate = new DateTime(2022, 10, 1);
            _reportGenerationRequest.ToDate = new DateTime(2022, 12, 2);
            _reportGenerationRequest.OutputType = OutputType.PDF;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/Attendance_ToPDF.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void Attendance_ToExcel()
        {
            _reportGenerationRequest.ReportType = ReportType.Attendance;
            _reportGenerationRequest.ReportData = _attendanceReportData;
            _reportGenerationRequest.FromDate = new DateTime(2022, 10, 1);
            _reportGenerationRequest.ToDate = new DateTime(2022, 12, 2);
            _reportGenerationRequest.OutputType = OutputType.Excel;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/Attendance_ToExcel.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void TermProgram_ToExcel()
        {
            _reportGenerationRequest.ReportType = ReportType.TermProgram;
            _reportGenerationRequest.ReportData = _termProgramReportData;
            _reportGenerationRequest.OutputType = OutputType.Excel;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults/TermProgram_ToExcel.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

    }
}
