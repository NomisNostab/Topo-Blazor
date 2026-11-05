using Newtonsoft.Json;
using Topo.Model.History;
using Topo.Model.Logbook;
using Topo.Model.Milestone;
using Topo.Model.OAS;
using Topo.Model.Progress;
using Topo.Model.ReportGeneration;
using TopoReportFunction;
using AdditionalAwardResult = Topo.Model.AdditionalAwards.Result;
using AdditionalAwardResults = Topo.Model.AdditionalAwards.GetUnitAchievementsResultModel;
using SiaResult = Topo.Model.SIA.Result;
using SiaResults = Topo.Model.SIA.GetSIAResultsModel;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class ScoutingHistory
    {
        private const string CubUnitId = "11111111-0000-0000-0000-000000000000";
        private const string ScoutUnitId = "22222222-0000-0000-0000-000000000000";

        private static ScoutingHistoryInput SampleInput()
        {
            var input = new ScoutingHistoryInput
            {
                MemberName = "Sample Member",
                AsOf = new DateTime(2026, 6, 1),
                CurrentSection = "scout",
                CurrentRole = "APL",
                UnitNames = new Dictionary<string, string> { [ScoutUnitId] = "Scout Unit" },
                AwardNames = new Dictionary<string, string> { ["aa-1"] = "Bronze Cross" }
            };

            input.IntroToScouting.results = new[]
            {
                new Topo.Model.Progress.Result { section = "cub", status = "awarded", imported = new Topo.Model.Progress.Imported { date_awarded = "2019-08-01" } }
            };
            input.IntroToSection.results = new[]
            {
                new Topo.Model.Progress.Result { section = "cub", status = "awarded", imported = new Topo.Model.Progress.Imported { date_awarded = "2019-09-01" } },
                new Topo.Model.Progress.Result { section = "scout", status = "awarded", imported = new Topo.Model.Progress.Imported { date_awarded = "2022-02-01" } }
            };

            input.Milestones.results = new[]
            {
                new MilestoneResult { section = "cub", status = "awarded", achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = 1 }, imported = new Topo.Model.Milestone.Imported { date_awarded = "2020-03-01" } },
                new MilestoneResult { section = "cub", status = "in_progress", achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = 2 } },
                new MilestoneResult { section = "scout", status = "awarded", achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = 2 }, status_updated = new DateTime(2023, 4, 4) }
            };

            input.Oas.results = new[]
            {
                OasResult("cub", "bushcraft", 1, "awarded", "2020-05-01"),
                OasResult("cub", "bushcraft", 2, "awarded", "2021-05-01"),
                OasResult("scout", "bushcraft", 3, "awarded", "2022-08-01"),
                OasResult("scout", "bushcraft", 7, "approved", "2025-01-01"),
                OasResult("scout", "camping", 1, "draft_review", "2025-02-01")
            };

            input.Sia.results = new[]
            {
                new SiaResult { section = "scout", status = "awarded", answers = new Topo.Model.SIA.Answers { special_interest_area_selection = "sia_adventure_sports", project_name = "Kayak trip" }, status_updated = new DateTime(2024, 1, 5) },
                new SiaResult { section = "scout", status = "draft_review", answers = new Topo.Model.SIA.Answers { special_interest_area_selection = "sia_stem_innovation" } }
            };

            input.AdditionalAwards.results = new[]
            {
                new AdditionalAwardResult { section = "scout", status = "awarded", achievement_meta = new Topo.Model.AdditionalAwards.Achievement_Meta { additional_award_id = "aa-1" }, status_updated = new DateTime(2024, 5, 8) }
            };

            input.AllAchievements.results = new[]
            {
                new MemberAchievement { id = "a", section = "cub", latest_submission = new LatestSubmission { unit_id = CubUnitId } },
                new MemberAchievement { id = "b", section = "scout", latest_submission = new LatestSubmission { unit_id = ScoutUnitId } },
                new MemberAchievement { id = "c", section = "scout" }
            };

            input.Logbook = new List<GetMemberLogbookDetailResultModel>
            {
                Entry("2020-05-01", "2020-05-03", new[] { "camping" }, 0, "Participant"),
                Entry("2022-06-10", "2022-06-12", new[] { "camping", "walking_hike" }, 8000, "Lead"),
                Entry("2024-03-01", "2024-03-01", new[] { "walking_hike" }, 4000, "assist"),
                Entry("2018-01-01", "2018-01-02", new[] { "camping" }, 0, "participant")
            };
            return input;
        }

        private static Topo.Model.OAS.Result OasResult(string section, string stream, int stage, string status, string awarded) =>
            new Topo.Model.OAS.Result
            {
                section = section,
                status = status,
                achievement_meta = new Topo.Model.OAS.Achievement_Meta { stream = stream, stage = stage },
                imported = new Topo.Model.OAS.Imported { date_awarded = awarded }
            };

        private static GetMemberLogbookDetailResultModel Entry(string start, string end, string[] categories, float walked, string role) =>
            new GetMemberLogbookDetailResultModel
            {
                start_date = start,
                end_date = end,
                categories = categories,
                distance_walkabout = walked,
                details = new Details { your_role = role }
            };

        [TestMethod]
        public void ReportsRenderAsPdfAndExcel()
        {
            var history = ScoutingHistoryBuilder.Build(SampleInput());
            SetEnvironment.SetSyncfusionLicense();
            var function = new Function();
            foreach (var (outputType, extension) in new[] { (OutputType.PDF, "pdf"), (OutputType.Excel, "xlsx") })
            {
                var request = new ReportGenerationRequest
                {
                    ReportType = ReportType.ScoutingHistory,
                    GroupName = "Sample Group",
                    Section = "scout",
                    UnitName = "Scout Unit",
                    OutputType = outputType,
                    ReportData = JsonConvert.SerializeObject(history)
                };
                var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) };
                var result = function.FunctionHandler(proxyRequest, null);
                Assert.IsFalse(string.IsNullOrEmpty(result), $"{extension} came back empty");
                File.WriteAllBytes($@"TestResults/ScoutingHistory.{extension}", Convert.FromBase64String(result));
            }
        }

        [TestMethod]
        public void SplitsBySectionInOrderWithDatesAndUnits()
        {
            var history = ScoutingHistoryBuilder.Build(SampleInput());

            Assert.AreEqual(2, history.Sections.Count);
            Assert.AreEqual("Cubs", history.Sections[0].SectionName);
            Assert.AreEqual("Cub unit (another group)", history.Sections[0].Units);
            Assert.AreEqual(new DateTime(2019, 9, 1), history.Sections[0].From);
            Assert.AreEqual(new DateTime(2022, 2, 1), history.Sections[0].To);
            Assert.AreEqual("2 years 5 months", history.Sections[0].TimeInSection);
            Assert.AreEqual("Scouts", history.Sections[1].SectionName);
            Assert.AreEqual("Scout Unit", history.Sections[1].Units);
            Assert.AreEqual(new DateTime(2026, 6, 1), history.Sections[1].To);
            Assert.AreEqual("APL", history.Sections[1].CurrentRole);
            Assert.AreEqual(new DateTime(2019, 8, 1), history.ScoutingStart);
            Assert.AreEqual("6 years 10 months", history.TimeInScouting);
        }

        [TestMethod]
        public void MilestonesAndFinalLevelPerSection()
        {
            var history = ScoutingHistoryBuilder.Build(SampleInput());

            Assert.AreEqual("Milestone 1", history.Sections[0].FinalMilestone);
            Assert.AreEqual("Milestone 2", history.Sections[1].FinalMilestone);
            Assert.AreEqual(2, history.Totals.Milestones); // the in progress milestone is not counted
        }

        [TestMethod]
        public void OasFinalLevelCountsEarlierSectionsAndProgressionsOnlyThisSection()
        {
            var history = ScoutingHistoryBuilder.Build(SampleInput());

            var cubBushcraft = history.Sections[0].Oas.Single();
            Assert.AreEqual(2, cubBushcraft.FinalStage);
            Assert.AreEqual(2, cubBushcraft.Progressions);

            // Scouts: stage 3 (1 progression) and stage 7 (counts as 2); camping is only in draft so is not counted
            var scoutBushcraft = history.Sections[1].Oas.Single();
            Assert.AreEqual("Bushcraft", scoutBushcraft.Stream);
            Assert.AreEqual(7, scoutBushcraft.FinalStage);
            Assert.AreEqual(3, scoutBushcraft.Progressions);
            Assert.AreEqual(3, history.Sections[1].OasProgressions);
            Assert.AreEqual(5, history.Totals.OasProgressions);
        }

        [TestMethod]
        public void SiaAndAdditionalAwards()
        {
            var history = ScoutingHistoryBuilder.Build(SampleInput());

            var scouts = history.Sections[1];
            Assert.AreEqual("Adventure Sports", scouts.Sia.Single().Area);
            Assert.AreEqual("Kayak trip", scouts.Sia.Single().Project);
            Assert.AreEqual(1, scouts.SiaInProgress);
            Assert.AreEqual("Bronze Cross", scouts.AdditionalAwards.Single().Name);
        }

        [TestMethod]
        public void LogbookEntriesGoToTheSectionTheMemberWasInOnThatDate()
        {
            var history = ScoutingHistoryBuilder.Build(SampleInput());

            var cubs = history.Sections[0];
            // the 2018 entry is before the Cub start, so it still belongs to the first section
            Assert.AreEqual(2, cubs.Camps);
            Assert.AreEqual(3, cubs.Nights);
            Assert.AreEqual(2, cubs.Participated);

            var scouts = history.Sections[1];
            Assert.AreEqual(1, scouts.Camps);
            Assert.AreEqual(2, scouts.Nights);
            Assert.AreEqual(2, scouts.Hikes);
            Assert.AreEqual(12.0, scouts.HikeKm);
            Assert.AreEqual(1, scouts.Led);
            Assert.AreEqual(1, scouts.Assisted);

            Assert.AreEqual(5, history.Totals.Nights);
            Assert.AreEqual(3, history.Totals.Camps);
        }

        [TestMethod]
        public void DurationIsWholeYearsAndMonths()
        {
            Assert.AreEqual("less than a month", ScoutingHistoryBuilder.Duration(new DateTime(2024, 1, 10), new DateTime(2024, 2, 9)));
            Assert.AreEqual("1 month", ScoutingHistoryBuilder.Duration(new DateTime(2024, 1, 10), new DateTime(2024, 2, 10)));
            Assert.AreEqual("1 year", ScoutingHistoryBuilder.Duration(new DateTime(2024, 1, 1), new DateTime(2025, 1, 1)));
            Assert.AreEqual("3 years 11 months", ScoutingHistoryBuilder.Duration(new DateTime(2020, 1, 15), new DateTime(2023, 12, 20)));
            Assert.AreEqual("", ScoutingHistoryBuilder.Duration(null, new DateTime(2024, 1, 1)));
        }
    }
}
