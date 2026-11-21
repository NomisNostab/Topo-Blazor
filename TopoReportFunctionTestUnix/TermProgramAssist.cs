using Syncfusion.XlsIO;
using Topo.Model.Program;
using Topo.Services;

namespace TopoReportFunctionTest
{
    // Assists are no longer a role in Terrain: the Assist column only shows when a listed event still has one
    [TestClass]
    public class TermProgramAssist
    {
        [TestInitialize]
        public void SetUp()
        {
            SetEnvironment.SetSyncfusionLicense();
        }

        private static EventListModel Event(string name, string assist = "") => new EventListModel
        {
            EventName = name,
            StartDateTime = new DateTime(2026, 11, 11, 19, 0, 0),
            EndDateTime = new DateTime(2026, 11, 11, 20, 30, 0),
            StartDateTimeDisplay = "11/11/26 19:00",
            EndDateTimeDisplay = "11/11/26 20:30",
            Lead = "Sam E",
            Assist = assist
        };

        private static List<string> Headings(List<EventListModel> events)
        {
            var sheet = new ReportService().GenerateTermProgramWorkbook(events, "Epping Scout Group", "scout", "Scout Unit", false).Worksheets[0];
            for (int row = 1; row <= 10; row++)
            {
                if (sheet.Range[row, 1].Text == "Start")
                    return Enumerable.Range(1, 16).Select(c => sheet.Range[row, c].Text).Where(t => !string.IsNullOrEmpty(t)).ToList();
            }
            Assert.Fail("No heading row found");
            return new List<string>();
        }

        [TestMethod]
        public void NoAssistColumnWhenNoEventHasAnAssist()
        {
            var headings = Headings(new List<EventListModel> { Event("Canoe Night"), Event("Raft making") });
            CollectionAssert.AreEqual(new[] { "Start", "End", "Date", "Week", "Start and Finish", "Meet at", "Activity", "Organiser", "Lead", "Duty Patrol", "Uniform", "Operoo/E1", "Challenge" }, headings);
        }

        [TestMethod]
        public void AssistColumnShowsWhenAnEventStillHasAnAssist()
        {
            var headings = Headings(new List<EventListModel> { Event("Canoe Night"), Event("Raft making", "Alex S") });
            CollectionAssert.AreEqual(new[] { "Start", "End", "Date", "Week", "Start and Finish", "Meet at", "Activity", "Organiser", "Lead", "Assist", "Duty Patrol", "Uniform", "Operoo/E1", "Challenge" }, headings);
        }

        [TestMethod]
        public void ThePageShowsTheAssistColumnOnlyWhenAnEventHasAnAssist()
        {
            var none = new ProgramPageViewModel { Events = new List<EventListModel> { Event("Canoe Night"), Event("Raft making", "  ") } };
            var some = new ProgramPageViewModel { Events = new List<EventListModel> { Event("Canoe Night"), Event("Raft making", "Alex S") } };
            Assert.IsFalse(none.ShowAssist);
            Assert.IsTrue(some.ShowAssist);
            Assert.IsFalse(new ProgramPageViewModel().ShowAssist);
        }
    }
}
