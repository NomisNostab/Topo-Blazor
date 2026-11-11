using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Topo.Model.ReportGeneration;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    // Reports for a unit with no youth members (an adults-only team), or with nothing to report yet
    [TestClass]
    public class AdultOnlyUnits
    {
        private static string Generate(ReportType type, OutputType output, string data)
        {
            SetEnvironment.SetSyncfusionLicense();
            var request = new ReportGenerationRequest
            {
                ReportType = type,
                GroupName = "Group",
                Section = "scout",
                UnitName = "Adventurous Activity Team",
                OutputType = output,
                ReportData = data
            };
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) };
            return new Function().FunctionHandler(proxyRequest, null);
        }

        private static string AttendanceWhere(Func<JToken, bool> keep)
        {
            var attendance = JObject.Parse(File.ReadAllText("ReportData/EventAttendance.json"));
            attendance["attendanceReportItems"] = new JArray(((JArray)attendance["attendanceReportItems"]!).Where(keep));
            return attendance.ToString(Formatting.None);
        }

        [DataTestMethod]
        [DataRow(OutputType.PDF)]
        [DataRow(OutputType.Excel)]
        public void AttendanceReportWithOnlyAdults(OutputType output)
        {
            var data = AttendanceWhere(item => (int?)item["IsAdultMember"] == 1);
            Assert.IsTrue(JObject.Parse(data)["attendanceReportItems"]!.Any(), "the sample has no adults");

            Assert.IsFalse(string.IsNullOrEmpty(Generate(ReportType.Attendance, output, data)));
        }

        [DataTestMethod]
        [DataRow(OutputType.PDF)]
        [DataRow(OutputType.Excel)]
        public void AttendanceReportWithOnlyYouth(OutputType output)
        {
            var data = AttendanceWhere(item => (int?)item["IsAdultMember"] == 0);

            Assert.IsFalse(string.IsNullOrEmpty(Generate(ReportType.Attendance, output, data)));
        }

        [DataTestMethod]
        [DataRow(OutputType.PDF)]
        [DataRow(OutputType.Excel)]
        public void AttendanceReportWithNoOneAtAll(OutputType output)
        {
            Assert.IsFalse(string.IsNullOrEmpty(Generate(ReportType.Attendance, output, "{}")));
        }

        [DataTestMethod]
        [DataRow(OutputType.PDF)]
        [DataRow(OutputType.Excel)]
        public void ApprovalsReportWithNoApprovals(OutputType output)
        {
            Assert.IsFalse(string.IsNullOrEmpty(Generate(ReportType.Approvals, output, "[]")));
        }

        [DataTestMethod]
        [DataRow(ReportType.MemberList)]
        [DataRow(ReportType.PatrolList)]
        [DataRow(ReportType.PatrolSheets)]
        [DataRow(ReportType.SignInSheet)]
        [DataRow(ReportType.Logbook)]
        [DataRow(ReportType.Wallchart)]
        [DataRow(ReportType.Milestone)]
        [DataRow(ReportType.SIA)]
        [DataRow(ReportType.TermProgram)]
        public void ListReportsWithNoMembers(ReportType type)
        {
            Assert.IsFalse(string.IsNullOrEmpty(Generate(type, OutputType.PDF, "[]")));
        }
    }
}
