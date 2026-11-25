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

        // A sheet for several sections puts each section on its own page, so it is not cut across pages
        [TestMethod]
        public void EachSectionStartsOnANewPage()
        {
            var members = new List<MemberListModel>
            {
                new MemberListModel { id = "1", first_name = "Cub", last_name = "One", section = "cub" },
                new MemberListModel { id = "2", first_name = "Scout", last_name = "Two", section = "scout" },
                new MemberListModel { id = "3", first_name = "Venturer", last_name = "Three", section = "venturer" },
                new MemberListModel { id = "4", first_name = "Leader", last_name = "Four", section = "scout", isAdultLeader = 1 }
            };
            var sheet = new ReportService().GenerateSignInSheetWorkbook(members, "Group", "", "Units", "Camp", new SignInSheetOptions()).Worksheets[0];
            // cubs, scouts, venturers and leaders: a break before each of the last three
            Assert.AreEqual(3, sheet.HPageBreaks.Count);
        }

        [TestMethod]
        public void OneSectionHasNoPageBreak()
        {
            var members = new List<MemberListModel>
            {
                new MemberListModel { id = "1", first_name = "Scout", last_name = "One", section = "scout" },
                new MemberListModel { id = "2", first_name = "Scout", last_name = "Two", section = "scout" }
            };
            var sheet = new ReportService().GenerateSignInSheetWorkbook(members, "Group", "scout", "Unit", "Camp", new SignInSheetOptions()).Worksheets[0];
            Assert.AreEqual(0, sheet.HPageBreaks.Count);
        }

        private static int PdfPages(List<MemberListModel> members, SignInSheetOptions options)
        {
            var workbook = new ReportService().GenerateSignInSheetWorkbook(members, "Group", "", "Units", "Camp", options);
            var renderer = new Syncfusion.XlsIORenderer.XlsIORenderer();
            using var pdf = renderer.ConvertToPDF(workbook);
            return pdf.Pages.Count;
        }

        private static List<MemberListModel> Section(string section, int count)
        {
            return Enumerable.Range(1, count).Select(i => new MemberListModel
            {
                id = section + i, first_name = section + " Member", last_name = "Number " + i, section = section, patrol_name = "Patrol " + (i % 4)
            }).ToList();
        }

        // Each section is on one page: a long section is shrunk to fit its page, not run on to a second one
        [TestMethod]
        public void ALongSectionStillFitsOnOnePage()
        {
            foreach (var count in new[] { 12, 30, 45 })
            {
                var members = Section("cub", 5).Concat(Section("scout", count)).Concat(Section("venturer", 3)).ToList();
                Assert.AreEqual(3, PdfPages(members, new SignInSheetOptions()), $"a section of {count}");
                Assert.AreEqual(3, PdfPages(members, new SignInSheetOptions { UseTimeInOutSignature = true, IncludeParticipationColumns = true }), $"a section of {count} with every column");
            }
        }

        // Portrait wherever it can be had: with the usual columns a sheet is portrait (a little shrunk to fit), not turned sideways
        [TestMethod]
        public void TheUsualSheetIsPortraitOnA4()
        {
            var members = JsonConvert.DeserializeObject<List<MemberListModel>>(File.ReadAllText("ReportData/SignInGroup.json"))!;
            var sheet = new ReportService().GenerateSignInSheetWorkbook(members, "Group", "scout", "Unit", "Camp", new SignInSheetOptions()).Worksheets[0];
            Assert.AreEqual(Syncfusion.XlsIO.ExcelPageOrientation.Portrait, sheet.PageSetup.Orientation);
            Assert.AreEqual(Syncfusion.XlsIO.ExcelPaperSize.PaperA4, sheet.PageSetup.PaperSize);
        }
    }
}
