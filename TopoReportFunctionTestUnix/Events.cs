using Newtonsoft.Json;
using Topo.Model.Members;
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
        public void SignInSheet_OrdersULThenByPatrolThenPLThenAPLThenAlphabetical()
        {
            MemberListModel Member(string id, string first, string patrol, string duty) =>
                new MemberListModel { id = id, first_name = first, last_name = "X", patrol_name = patrol, patrol_duty = duty, section = "scout" };
            var members = new List<MemberListModel>
            {
                Member("1", "Zoe", "Hawks", ""),
                Member("2", "Ben", "Eagles", ""),
                Member("3", "Cal", "Eagles", "APL"),
                Member("4", "Amy", "Eagles", ""),
                Member("5", "Dan", "Eagles", "PL"),
                Member("6", "Ava", "", ""),
                Member("7", "Kit", "Hawks", "PL"),
                Member("8", "Zed", "", "UL"),
                Member("9", "Una", "Reds", ""),
                Member("10", "Tom", "Reds", "Seconder"),
                Member("11", "Sam", "Reds", "Sixer"),
            };

            var names = SignInSheetLayout.GetDisplayNames(members, false);
            var ordered = SignInSheetLayout.GetGroups(members, names).Single().Members.Select(m => m.first_name);

            CollectionAssert.AreEqual(new[] { "Zed", "Dan", "Cal", "Amy", "Ben", "Kit", "Zoe", "Sam", "Tom", "Una", "Ava" }, ordered.ToList());
        }

        [TestMethod]
        public void SignInSheet_Group_Portrait_ToPDF()
        {
            _reportGenerationRequest.ReportData = File.ReadAllText("ReportData/SignInGroup.json");
            _reportGenerationRequest.EventName = "Group Camp";
            _reportGenerationRequest.SignInSheetOptions.UseTimeInOutSignature = false;
            _reportGenerationRequest.SignInSheetOptions.IncludeParticipationColumns = false;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            File.WriteAllBytes(@"TestResults/SignInSheet_Group_Portrait_ToPDF.pdf", Convert.FromBase64String(functionResult));
        }

        [TestMethod]
        public void SignInSheet_FullNamesThatMatchGetNumbered()
        {
            MemberListModel Member(string id, string first, string last) =>
                new MemberListModel { id = id, first_name = first, last_name = last, section = "scout" };
            var members = new List<MemberListModel> { Member("1", "Ethan", "Wu"), Member("2", "Ethan", "Wu"), Member("3", "Mia", "Chen") };

            var names = SignInSheetLayout.GetDisplayNames(members, true);

            Assert.AreEqual("Ethan Wu (1)", names["1"]);
            Assert.AreEqual("Ethan Wu (2)", names["2"]);
            Assert.AreEqual("Mia Chen", names["3"]);
        }

        [TestMethod]
        public void SignInSheet_TheTitleRowIsTallEnoughForALongEventName()
        {
            SetEnvironment.SetSyncfusionLicense();
            var members = JsonConvert.DeserializeObject<List<Topo.Model.Members.MemberListModel>>(File.ReadAllText("ReportData/SignInGroup.json"))!;
            var options = new Topo.Model.Program.SignInSheetOptions { UseTimeInOutSignature = false, IncludeParticipationColumns = false };
            var service = new Topo.Services.ReportService();

            double TitleRowHeight(string eventName) =>
                service.GenerateSignInSheetWorkbook(members, "Group", "scout", "Unit", eventName, options).Worksheets[0].GetRowHeight(3);

            Assert.AreEqual(30, TitleRowHeight("Group Camp")); // a short name keeps the usual height
            Assert.IsTrue(TitleRowHeight(new string('x', 110)) >= 60, "a long name needs more than one line");
            Assert.IsTrue(TitleRowHeight(new string('x', 200)) > TitleRowHeight(new string('x', 110)), "a longer name needs more lines");
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void SignInSheet_ALongEventNameIsNotCutOff(bool timeInOut)
        {
            _reportGenerationRequest.ReportData = File.ReadAllText("ReportData/SignInGroup.json");
            _reportGenerationRequest.EventName = "Combined District Cub Scout and Scout Leaders' Planning Weekend at the Mountain Retreat for the Whole Region ENDWORD";
            _reportGenerationRequest.SignInSheetOptions.UseTimeInOutSignature = timeInOut;
            _reportGenerationRequest.SignInSheetOptions.IncludeParticipationColumns = false;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);

            var functionResult = _function.FunctionHandler(proxyRequest, null);

            Assert.IsFalse(string.IsNullOrEmpty(functionResult));
            File.WriteAllBytes($@"TestResults/SignInSheet_LongName_{(timeInOut ? "TimeInOut" : "Tick")}.pdf", Convert.FromBase64String(functionResult));
            using var document = new Syncfusion.Pdf.Parsing.PdfLoadedDocument(Convert.FromBase64String(functionResult));
            var text = document.Pages[0].ExtractText().Replace("\r", " ").Replace("\n", " ");
            // the whole name is there, however it wraps (the first and last words are what a cut-off loses)
            StringAssert.Contains(text, "Combined");
            StringAssert.Contains(text, "ENDWORD");
        }

        [TestMethod]
        public void SignInSheet_Group_NoEvent_ToPDF()
        {
            _reportGenerationRequest.ReportData = File.ReadAllText("ReportData/SignInGroup.json");
            _reportGenerationRequest.EventName = "";
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            File.WriteAllBytes(@"TestResults/SignInSheet_Group_NoEvent_ToPDF.pdf", Convert.FromBase64String(functionResult));
        }

        [TestMethod]
        public void SignInSheet_Group_ToExcel()
        {
            _reportGenerationRequest.ReportData = File.ReadAllText("ReportData/SignInGroup.json");
            _reportGenerationRequest.EventName = "Group Camp";
            _reportGenerationRequest.OutputType = OutputType.Excel;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            File.WriteAllBytes(@"TestResults/SignInSheet_Group_ToExcel.xlsx", Convert.FromBase64String(functionResult));
        }

        [TestMethod]
        public void SignInSheet_Group_TimeInOut_ToPDF()
        {
            _reportGenerationRequest.ReportData = File.ReadAllText("ReportData/SignInGroup.json");
            _reportGenerationRequest.EventName = "Group Camp";
            _reportGenerationRequest.SignInSheetOptions.UseTimeInOutSignature = true;
            _reportGenerationRequest.SignInSheetOptions.IncludeRegisteredColumn = true;
            _reportGenerationRequest.SignInSheetOptions.IncludePaidColumn = true;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            File.WriteAllBytes(@"TestResults/SignInSheet_Group_TimeInOut_ToPDF.pdf", Convert.FromBase64String(functionResult));
        }

        [TestMethod]
        public void SignInSheet_Group_TickBox_FullNames_ToPDF()
        {
            _reportGenerationRequest.ReportData = File.ReadAllText("ReportData/SignInGroup.json");
            _reportGenerationRequest.EventName = "Group Camp";
            _reportGenerationRequest.SignInSheetOptions.UseTimeInOutSignature = false;
            _reportGenerationRequest.SignInSheetOptions.UseFullNames = true;
            _reportGenerationRequest.SignInSheetOptions.IncludeParticipationColumns = false;
            _reportGenerationRequest.SignInSheetOptions.IncludePaidColumn = true;
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            File.WriteAllBytes(@"TestResults/SignInSheet_Group_TickBox_FullNames_ToPDF.pdf", Convert.FromBase64String(functionResult));
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
