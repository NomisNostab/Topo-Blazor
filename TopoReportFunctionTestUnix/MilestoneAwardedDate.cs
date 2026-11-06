using Topo.Model.Milestone;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class MilestoneAwardedDate
    {
        [TestMethod]
        public void AnImportedMilestoneUsesTheDateItWasAwarded()
        {
            var result = new MilestoneResult
            {
                status_updated = new DateTime(2025, 11, 3, 14, 20, 0),
                imported = new Imported { date_awarded = "2023-03-18" }
            };

            Assert.AreEqual(new DateTime(2023, 3, 18), MilestoneDates.AwardedDate(result));
        }

        [TestMethod]
        public void AMilestoneAwardedInTerrainUsesWhenItsStatusLastChanged()
        {
            var updated = new DateTime(2025, 11, 3, 14, 20, 0);

            Assert.AreEqual(updated, MilestoneDates.AwardedDate(new MilestoneResult { status_updated = updated }));
            Assert.AreEqual(updated, MilestoneDates.AwardedDate(new MilestoneResult { status_updated = updated, imported = new Imported { date_awarded = "not a date" } }));
        }
    }
}
