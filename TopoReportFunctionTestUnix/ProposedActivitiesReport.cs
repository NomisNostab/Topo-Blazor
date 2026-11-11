using Newtonsoft.Json;
using Syncfusion.XlsIO;
using Topo.Model.Program;
using Topo.Model.ReportGeneration;
using Topo.Services;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class ProposedActivitiesReport
    {
        [TestInitialize]
        public void SetUp()
        {
            SetEnvironment.SetSyncfusionLicense();
        }

        private static ProposedActivitiesReportData Data(DateTime? since = null, bool known = true) => new ProposedActivitiesReportData
        {
            Since = since,
            Proposals = new List<ProposedActivityModel>
            {
                new ProposedActivityModel
                {
                    Id = "1", Title = "Capture The Flag", ChallengeAreaKey = "outdoors", Submitted = new DateTime(2026, 3, 10), ProposedBy = "Sam Example", Patrol = "Lyrebird",
                    SuggestedLead = "The leaders", Description = "Two teams and flags", Justification = "Everyone likes it", TimesProposed = 2,
                    TimesRun = known ? 3 : (int?)null, LastRun = known ? new DateTime(2026, 5, 20) : (DateTime?)null
                },
                new ProposedActivityModel
                {
                    Id = "2", Title = "Origami", ChallengeAreaKey = "creative", Submitted = new DateTime(2025, 9, 14), ProposedBy = "Alex Sample",
                    TimesProposed = 1, TimesRun = known ? 0 : (int?)null
                }
            }
        };

        private static IWorksheet Sheet(ProposedActivitiesReportData data)
        {
            var workbook = new ReportService().GenerateProposedActivitiesWorkbook(data, "Epping Scout Group", "scout", "Scout Unit");
            return workbook.Worksheets[0];
        }

        [TestMethod]
        public void TheSheetHasTheHeadingsAndOneRowToAProposal()
        {
            var sheet = Sheet(Data());
            Assert.AreEqual("Epping Scout Group", sheet.Range[1, 2].Text);
            Assert.AreEqual("Scout Unit", sheet.Range[2, 2].Text);
            Assert.AreEqual("Proposed Activities", sheet.Range[3, 2].Text);
            CollectionAssert.AreEqual(new[] { "Activity", "Challenge Area", "Date Submitted", "Proposed By", "Suggested Lead", "Times Proposed", "Times Run", "Last Run", "What", "Why" },
                Enumerable.Range(1, 10).Select(c => sheet.Range[6, c].Text).ToList());

            CollectionAssert.AreEqual(new[] { "Capture The Flag", "Outdoors", "10/03/2026", "Sam Example, Lyrebird", "The leaders", "2", "3", "20/05/2026", "Two teams and flags", "Everyone likes it" },
                Enumerable.Range(1, 10).Select(c => sheet.Range[7, c].DisplayText).ToList());
            CollectionAssert.AreEqual(new[] { "Origami", "Creative", "14/09/2025", "Alex Sample", "", "1", "0", "", "", "" },
                Enumerable.Range(1, 10).Select(c => sheet.Range[8, c].DisplayText).ToList());
            Assert.AreEqual("", sheet.Range[9, 1].DisplayText);
        }

        [TestMethod]
        public void TheNoteSaysWhereTheRunCountsStarted()
        {
            StringAssert.Contains(Sheet(Data()).Range[4, 2].Text, "all events up to today");
            StringAssert.Contains(Sheet(Data(new DateTime(2025, 7, 1))).Range[4, 2].Text, "from 01/07/2025");
        }

        [TestMethod]
        public void UnknownRunsAreLeftBlankAndSaidSo()
        {
            var sheet = Sheet(Data(known: false));
            StringAssert.Contains(sheet.Range[4, 2].Text, "could not be read");
            Assert.AreEqual("", sheet.Range[7, 7].DisplayText);
            Assert.AreEqual("", sheet.Range[8, 7].DisplayText);
            Assert.AreEqual("2", sheet.Range[7, 6].DisplayText);
        }

        [TestMethod]
        public void NoProposalsGivesJustTheHeadings()
        {
            var sheet = Sheet(new ProposedActivitiesReportData());
            Assert.AreEqual("Activity", sheet.Range[6, 1].Text);
            Assert.AreEqual("", sheet.Range[7, 1].DisplayText);
        }

        [TestMethod]
        public void TheFunctionMakesAPdfAndASpreadsheet()
        {
            var function = new Function();
            foreach (var (output, start) in new[] { (OutputType.PDF, "%PDF"), (OutputType.Excel, "PK") })
            {
                var request = new ReportGenerationRequest
                {
                    ReportType = ReportType.ProposedActivities, GroupName = "Epping Scout Group", Section = "scout", UnitName = "Scout Unit",
                    OutputType = output, ReportData = JsonConvert.SerializeObject(Data())
                };
                var result = function.FunctionHandler(new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) }, null!);
                var bytes = Convert.FromBase64String(result);
                Assert.IsTrue(bytes.Length > 1000, output.ToString());
                Assert.AreEqual(start, System.Text.Encoding.ASCII.GetString(bytes, 0, start.Length), output.ToString());
                // kept to look at, like the other report tests
                Directory.CreateDirectory("TestResults");
                File.WriteAllBytes(output == OutputType.PDF ? "TestResults/ProposedActivities_ToPDF.pdf" : "TestResults/ProposedActivities_ToExcel.xlsx", bytes);
            }
        }

        [TestMethod]
        public void TheReportDataSurvivesTheJourneyToTheFunction()
        {
            var back = JsonConvert.DeserializeObject<ProposedActivitiesReportData>(JsonConvert.SerializeObject(Data(new DateTime(2025, 7, 1))))!;
            Assert.AreEqual(new DateTime(2025, 7, 1), back.Since);
            Assert.AreEqual(2, back.Proposals.Count);
            Assert.AreEqual("Outdoors", back.Proposals[0].ChallengeArea);
            Assert.AreEqual(3, back.Proposals[0].TimesRun);
            Assert.AreEqual("Sam Example", back.Proposals[0].ProposedBy);
        }
    }
}
