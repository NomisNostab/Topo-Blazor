using Topo.Model.Milestone;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class MilestoneLayoutTests
    {
        private static MilestoneSummaryListModel Scout(string name, string patrol, int level, bool milestone3Awarded = false) => new MilestoneSummaryListModel
        {
            memberName = name, memberPatrol = patrol, currentLevel = level, milestone3Awarded = milestone3Awarded, percentComplete = 50,
            milestone1ParticipateCommunity = 1, milestone1ParticipateOutdoors = 2, milestone1ParticipateCreative = 3, milestone1ParticipatePersonalGrowth = 4, milestone1Lead = 1, milestone1LeadArea = "Co",
            milestone2ParticipateCommunity = 5, milestone2ParticipateOutdoors = 0, milestone2ParticipateCreative = 2, milestone2ParticipatePersonalGrowth = 1, milestone2Lead = 3, milestone2LeadArea = "O-3",
            milestone3ParticipateCommunity = 4, milestone3ParticipateOutdoors = 3, milestone3ParticipateCreative = 2, milestone3ParticipatePersonalGrowth = 1, milestone3Lead = 6, milestone3LeadArea = "Co-2 O-4"
        };

        private static (int r, int g, int b) Rgb(string hex) => (Convert.ToInt32(hex.Substring(1, 2), 16), Convert.ToInt32(hex.Substring(3, 2), 16), Convert.ToInt32(hex.Substring(5, 2), 16));

        [TestMethod]
        public void TheShadesAreTheOnesTheReportAlwaysUsed()
        {
            // the colours as the report had them before they moved into the model, as red, green and blue
            var scale1 = new[] { (236, 82, 82), (156, 242, 80) };
            var scale2 = new[] { (236, 82, 82), (253, 238, 0), (156, 242, 80) };
            var scale4 = new[] { (236, 82, 82), (244, 157, 43), (253, 238, 0), (206, 240, 38), (156, 242, 80) };
            var scale5 = new[] { (236, 82, 82), (241, 131, 57), (248, 189, 26), (221, 239, 26), (184, 241, 57), (156, 242, 80) };
            var scale6 = new[] { (236, 82, 82), (241, 129, 57), (248, 194, 23), (253, 238, 0), (218, 240, 29), (183, 241, 58), (156, 242, 80) };

            CollectionAssert.AreEqual(scale1, MilestoneColours.Scale1.Select(Rgb).ToArray());
            CollectionAssert.AreEqual(scale2, MilestoneColours.Scale2.Select(Rgb).ToArray());
            CollectionAssert.AreEqual(scale4, MilestoneColours.Scale4.Select(Rgb).ToArray());
            CollectionAssert.AreEqual(scale5, MilestoneColours.Scale5.Select(Rgb).ToArray());
            CollectionAssert.AreEqual(scale6, MilestoneColours.Scale6.Select(Rgb).ToArray());
        }

        [TestMethod]
        public void EachMilestoneShadesParticipateAndLeadAgainstItsOwnTargets()
        {
            Assert.AreEqual(MilestoneColours.Scale6[3], MilestoneColours.For(1, true, 3));  // 6 to participate in
            Assert.AreEqual(MilestoneColours.Scale2[1], MilestoneColours.For(1, false, 1)); // 2 to lead
            Assert.AreEqual(MilestoneColours.Scale5[2], MilestoneColours.For(2, true, 2));
            Assert.AreEqual(MilestoneColours.Scale4[2], MilestoneColours.For(2, false, 2));
            Assert.AreEqual(MilestoneColours.Scale4[1], MilestoneColours.For(3, true, 1));
            Assert.AreEqual(MilestoneColours.Scale6[5], MilestoneColours.For(3, false, 5));

            // more than the target is as green as meeting it
            Assert.AreEqual(MilestoneColours.Scale6[6], MilestoneColours.For(1, true, 9));
            Assert.AreEqual(MilestoneColours.Scale2[2], MilestoneColours.For(1, false, 40));
            // nothing done is red; a skipped milestone (a negative count) and an unknown one are white
            Assert.AreEqual(MilestoneColours.Scale4[0], MilestoneColours.For(3, true, 0));
            Assert.AreEqual(MilestoneColours.White, MilestoneColours.For(2, true, -1));
            Assert.AreEqual(MilestoneColours.White, MilestoneColours.For(4, true, 3));
        }

        [TestMethod]
        public void ScoutsAreSectionedByMilestoneAndThoseWhoFinishedMilestone3GetTheirOwn()
        {
            var groups = MilestoneLayout.Build(new[]
            {
                Scout("Zoe", "Red", 3),
                Scout("Amy", "Red", 1),
                Scout("Ben", "Blue", 3, milestone3Awarded: true),
                Scout("Cat", "Blue", 2)
            }, breakByPatrol: false);

            var unit = groups.Single();
            Assert.AreEqual("", unit.Patrol);
            CollectionAssert.AreEqual(new[] { "Milestone 1", "Milestone 2", "Milestone 3", "Milestone 3 - Completed" }, unit.Sections.Select(s => s.Title).ToList());
            CollectionAssert.AreEqual(new[] { "Participate - 6", "Participate - 5", "Participate - 4", "Participate - 4" }, unit.Sections.Select(s => s.ParticipateHeading).ToList());
            CollectionAssert.AreEqual(new[] { "Lead - 2", "Lead - 4", "Lead - 6", "Lead - 6" }, unit.Sections.Select(s => s.LeadHeading).ToList());
            Assert.AreEqual("Ben", unit.Sections.Last().Rows.Single().Name);
        }

        [TestMethod]
        public void BreakingByPatrolGivesEachPatrolItsOwnSheetInOrder()
        {
            var groups = MilestoneLayout.Build(new[] { Scout("Zoe", "Red", 1), Scout("Amy", "Red", 1), Scout("Ben", "Blue", 2), Scout("Pat", "", 1) }, breakByPatrol: true);

            CollectionAssert.AreEqual(new[] { "", "Blue", "Red" }, groups.Select(g => g.Patrol).ToList());
            CollectionAssert.AreEqual(new[] { "Amy", "Zoe" }, groups.Single(g => g.Patrol == "Red").Sections.Single().Rows.Select(r => r.Name).ToList());
        }

        [TestMethod]
        public void EachScoutsCountsComeFromTheirMilestonesOwnFieldsWithTheShadeForThatMilestone()
        {
            var groups = MilestoneLayout.Build(new[] { Scout("Amy", "Red", 2) }, breakByPatrol: false);
            var row = groups.Single().Sections.Single().Rows.Single();

            CollectionAssert.AreEqual(new[] { 5, 0, 2, 1 }, row.Participate.Select(c => c.Count).ToList());
            Assert.AreEqual(MilestoneColours.Scale5[5], row.Participate[0].Colour);
            Assert.AreEqual(MilestoneColours.Scale5[0], row.Participate[1].Colour);
            Assert.AreEqual(3, row.Lead.Count);
            Assert.AreEqual(MilestoneColours.Scale4[3], row.Lead.Colour);
            Assert.AreEqual("O-3", row.LeadAreas);
            Assert.AreEqual(50, row.PercentComplete);
        }

        [TestMethod]
        public void TheAverageRowAveragesEachCountColumn()
        {
            var amy = Scout("Amy", "Red", 1);
            var zoe = Scout("Zoe", "Red", 1);
            zoe.milestone1ParticipateCommunity = 4;
            zoe.milestone1Lead = 2;

            var averages = MilestoneLayout.Build(new[] { amy, zoe }, false).Single().Sections.Single().Averages();

            CollectionAssert.AreEqual(new[] { 2.5, 2.0, 3.0, 4.0, 1.5 }, averages);
        }

        [TestMethod]
        public void NoScoutsGivesNothingToLayOut()
        {
            Assert.AreEqual(0, MilestoneLayout.Build(new List<MilestoneSummaryListModel>(), false).Count);
            Assert.AreEqual(0, MilestoneLayout.Build(new List<MilestoneSummaryListModel>(), true).Count);
            Assert.AreEqual(0, new MilestoneSection().Averages().Sum()); // an empty section averages nothing rather than failing
        }
    }
}
