using Topo.Model.Milestone;
using Topo.Model.Progress;
using OasResult = Topo.Model.OAS.Result;

namespace TopoReportFunctionTest
{
    // Terrain can hold several records for one achievement; the one used must not depend on the order they come back in
    [TestClass]
    public class AchievementSelectionTests
    {
        private static Topo.Model.Progress.Result Intro(string section, string status, DateTime entered) =>
            new Topo.Model.Progress.Result { section = section, status = status, status_updated = entered };

        private static MilestoneResult Milestone(int stage, string status, DateTime entered, string section = "scout") =>
            new MilestoneResult { section = section, status = status, status_updated = entered, achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = stage } };

        private static OasResult Oas(string stream, int stage, string status = "awarded", string section = "scout") =>
            new OasResult { section = section, status = status, achievement_meta = new Topo.Model.OAS.Achievement_Meta { stream = stream, stage = stage } };

        [TestMethod]
        public void IntroUsesTheMostRecentlyEnteredAwardedRecordInEitherOrder()
        {
            var early = Intro("scout", "awarded", new DateTime(2021, 1, 1));
            var late = Intro("scout", "awarded", new DateTime(2024, 1, 1));

            Assert.AreSame(late, AchievementSelection.LatestAwardedIntro(new[] { early, late }));
            Assert.AreSame(late, AchievementSelection.LatestAwardedIntro(new[] { late, early }));
        }

        [TestMethod]
        public void IntroIgnoresOtherStatusesAndOtherSections()
        {
            var awarded = Intro("scout", "awarded", new DateTime(2021, 1, 1));
            var draft = Intro("scout", "in_progress", new DateTime(2024, 1, 1));
            var cub = Intro("cub", "awarded", new DateTime(2025, 1, 1));

            Assert.AreSame(awarded, AchievementSelection.LatestAwardedIntro(new[] { draft, cub, awarded }, "SCOUT"));
            Assert.AreSame(cub, AchievementSelection.LatestAwardedIntro(new[] { draft, cub, awarded })); // no section given: any section
            Assert.IsNull(AchievementSelection.LatestAwardedIntro(new[] { draft }, "scout"));
            Assert.IsNull(AchievementSelection.LatestAwardedIntro(new Topo.Model.Progress.Result[0]));
        }

        [TestMethod]
        public void PeakAwardUsesTheMostRecentlyEnteredAwardedRecordInTheSection()
        {
            PeakAwardResult Peak(string section, string status, DateTime entered) => new PeakAwardResult { section = section, status = status, status_updated = entered };
            var early = Peak("scout", "awarded", new DateTime(2022, 1, 1));
            var late = Peak("scout", "awarded", new DateTime(2023, 1, 1));
            var other = Peak("venturer", "awarded", new DateTime(2025, 1, 1));

            Assert.AreSame(late, AchievementSelection.LatestAwardedPeak(new[] { late, early, other }, "scout"));
            Assert.AreSame(late, AchievementSelection.LatestAwardedPeak(new[] { early, other, late }, "scout"));
            Assert.IsNull(AchievementSelection.LatestAwardedPeak(new PeakAwardResult[0], "scout"));
        }

        [TestMethod]
        public void AnAwardedMilestoneBeatsALaterInProgressDuplicateInEitherOrder()
        {
            var awarded = Milestone(1, "awarded", new DateTime(2023, 1, 1));
            var inProgress = Milestone(1, "in_progress", new DateTime(2024, 1, 1));

            Assert.AreSame(awarded, AchievementSelection.OnePerMilestoneStage(new[] { awarded, inProgress }).Single());
            Assert.AreSame(awarded, AchievementSelection.OnePerMilestoneStage(new[] { inProgress, awarded }).Single());
        }

        [TestMethod]
        public void MilestoneKeepsTheLatestOfEachKindAndOneRecordPerStageInStageOrder()
        {
            var awardedEarly = Milestone(2, "awarded", new DateTime(2022, 1, 1));
            var awardedLate = Milestone(2, "awarded", new DateTime(2023, 1, 1));
            var stageOne = Milestone(1, "awarded", new DateTime(2021, 1, 1));
            var draftEarly = Milestone(3, "in_progress", new DateTime(2022, 1, 1));
            var draftLate = Milestone(3, "in_progress", new DateTime(2024, 1, 1));

            var chosen = AchievementSelection.OnePerMilestoneStage(new[] { draftEarly, awardedEarly, draftLate, awardedLate, stageOne });

            CollectionAssert.AreEqual(new[] { stageOne, awardedLate, draftLate }, chosen);
            Assert.AreEqual(0, AchievementSelection.OnePerMilestoneStage(new MilestoneResult[0]).Count);
        }

        [TestMethod]
        public void MilestonesAwardedInDifferentSectionsAreNotDuplicatesOfEachOther()
        {
            // the same member earned Milestone 1 as a Cub and again as a Scout, and Milestone 2 as a Scout
            var cubOne = Milestone(1, "awarded", new DateTime(2020, 1, 1), section: "cub");
            var scoutOne = Milestone(1, "awarded", new DateTime(2023, 1, 1), section: "scout");
            var scoutTwo = Milestone(2, "awarded", new DateTime(2024, 1, 1), section: "scout");

            var chosen = AchievementSelection.OnePerMilestoneStage(new[] { scoutTwo, scoutOne, cubOne });

            CollectionAssert.AreEqual(new[] { cubOne, scoutOne, scoutTwo }, chosen);
        }

        [TestMethod]
        public void DuplicatesAreStillResolvedWithinEachSectionWhenSectionsAreMixed()
        {
            var cubAwarded = Milestone(1, "awarded", new DateTime(2020, 1, 1), section: "cub");
            var cubDraft = Milestone(1, "in_progress", new DateTime(2022, 1, 1), section: "cub");
            var scoutDraft = Milestone(1, "in_progress", new DateTime(2023, 1, 1), section: "Scout");
            var scoutAwarded = Milestone(1, "awarded", new DateTime(2024, 1, 1), section: "scout");

            var chosen = AchievementSelection.OnePerMilestoneStage(new[] { cubDraft, scoutDraft, cubAwarded, scoutAwarded });

            CollectionAssert.AreEqual(new[] { cubAwarded, scoutAwarded }, chosen);
        }

        [TestMethod]
        public void OasStageAwardedTwiceCountsOnce()
        {
            var results = new[] { Oas("bushcraft", 1), Oas("bushcraft", 1), Oas("bushcraft", 2) };

            Assert.AreEqual(2, AchievementSelection.OasProgressions(results, "scout", includeApproved: false));
        }

        [TestMethod]
        public void OasStagesSevenAndAboveCountTwiceAndOnlyInTheSection()
        {
            var results = new[]
            {
                Oas("camping", 3), Oas("camping", 7), Oas("camping", 7), Oas("hiking", 8),
                Oas("bushcraft", 4, section: "cub"), // another section
                Oas("alpine", 1, status: "draft_review")
            };

            Assert.AreEqual(1 + 2 + 2, AchievementSelection.OasProgressions(results, "Scout", includeApproved: false));
        }

        [TestMethod]
        public void OasApprovedCountsOnlyWhenAsked()
        {
            var results = new[] { Oas("camping", 1), Oas("camping", 2, status: "approved") };

            Assert.AreEqual(1, AchievementSelection.OasProgressions(results, "scout", includeApproved: false));
            Assert.AreEqual(2, AchievementSelection.OasProgressions(results, "scout", includeApproved: true));
            Assert.AreEqual(0, AchievementSelection.OasProgressions(new OasResult[0], "scout", includeApproved: true));
        }
    }
}
