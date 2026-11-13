using Newtonsoft.Json;
using Topo.Model.Members;
using Topo.Model.Program;
using Topo.Services;

namespace TopoReportFunctionTest
{
    // Assists are no longer a role in Terrain, so the sign in sheet's tick columns are Participate and Lead only
    [TestClass]
    public class SignInSheetParticipation
    {
        [TestInitialize]
        public void SetUp()
        {
            SetEnvironment.SetSyncfusionLicense();
        }

        private static List<(int Row, int Column, string Text)> HeadingCells(bool includeParticipation)
        {
            var members = JsonConvert.DeserializeObject<List<MemberListModel>>(File.ReadAllText("ReportData/SignInGroup.json"))!;
            var options = new SignInSheetOptions { UseTimeInOutSignature = false, IncludeParticipationColumns = includeParticipation };
            var sheet = new ReportService().GenerateSignInSheetWorkbook(members, "Group", "scout", "Unit", "Camp", options).Worksheets[0];
            var cells = new List<(int, int, string)>();
            for (var row = 1; row <= 12; row++)
                for (var column = 1; column <= 20; column++)
                    if (!string.IsNullOrEmpty(sheet.Range[row, column].Text))
                        cells.Add((row, column, sheet.Range[row, column].Text.Trim()));
            return cells;
        }

        [TestMethod]
        public void TheTickColumnsAreParticipateAndLeadWithNoAssist()
        {
            var cells = HeadingCells(true);
            Assert.IsFalse(cells.Any(c => c.Text == "Assist"), "there should be no Assist column");
            // the heading row repeats for each group of members, so look at the first one
            var participate = cells.First(c => c.Text == "Participate");
            var lead = cells.First(c => c.Row == participate.Row && c.Text == "Lead");
            Assert.AreEqual(participate.Column + 1, lead.Column, "Lead sits straight after Participate");
        }

        [TestMethod]
        public void WithoutTheTickColumnsThereIsNeitherParticipateNorLead()
        {
            var cells = HeadingCells(false);
            Assert.IsFalse(cells.Any(c => c.Text is "Participate" or "Assist" or "Lead"));
        }
    }
}
