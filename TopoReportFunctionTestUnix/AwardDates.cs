using Newtonsoft.Json;
using Topo.Model;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class AwardDates
    {
        private static readonly DateTime StatusUpdated = new DateTime(2026, 3, 20, 9, 30, 0);

        [TestMethod]
        public void TerrainsOwnDateWinsOverTheImportedDateAndWhenItWasLastUpdated()
        {
            var date = AwardDate.Resolve("2025-06-01", "2024-01-15", StatusUpdated);
            Assert.AreEqual(new DateTime(2025, 6, 1), date);
        }

        [TestMethod]
        public void TheImportedDateIsUsedWhenThereIsNoDateOfTerrainsOwn()
        {
            Assert.AreEqual(new DateTime(2024, 1, 15), AwardDate.Resolve("", "2024-01-15", StatusUpdated));
            Assert.AreEqual(new DateTime(2024, 1, 15), AwardDate.Resolve(null, "2024-01-15", StatusUpdated));
        }

        [TestMethod]
        public void WhenTerrainHasNeitherDateItFallsBackToWhenTheRecordLastChanged()
        {
            Assert.AreEqual(StatusUpdated, AwardDate.Resolve("", "", StatusUpdated));
            Assert.AreEqual(StatusUpdated, AwardDate.Resolve(null, null, StatusUpdated));
        }

        [TestMethod]
        public void ADateThatIsNotYyyyMmDdCountsAsMissingInsteadOfCrashing()
        {
            Assert.AreEqual(new DateTime(2024, 1, 15), AwardDate.Resolve("not a date", "2024-01-15", StatusUpdated));
            Assert.AreEqual(StatusUpdated, AwardDate.Resolve("13/02/2025", "", StatusUpdated));
        }

        [TestMethod]
        public void RecordedIsNullWhenOnlyStatusUpdatedExists()
        {
            Assert.IsNull(AwardDate.Recorded("", ""));
            Assert.AreEqual(new DateTime(2025, 6, 1), AwardDate.Recorded("2025-06-01", ""));
        }

        // The reason this exists: the record models used to have date_awarded only inside "imported", so the top-level
        // one was dropped when Terrain's JSON was read. Each kind of record must keep it now.
        private const string Awarded = "{\"results\":[{\"status\":\"awarded\",\"status_updated\":\"2026-03-20T09:30:00+11:00\",\"date_awarded\":\"2025-06-01\"}]}";

        [TestMethod]
        public void EveryKindOfAwardKeepsTerrainsOwnDateWhenReadFromJson()
        {
            Assert.AreEqual("2025-06-01", JsonConvert.DeserializeObject<Topo.Model.Milestone.GetMilestoneResultModel>(Awarded)!.results[0].date_awarded, "milestone");
            Assert.AreEqual("2025-06-01", JsonConvert.DeserializeObject<Topo.Model.OAS.GetUnitAchievementsResultsModel>(Awarded)!.results[0].date_awarded, "OAS");
            Assert.AreEqual("2025-06-01", JsonConvert.DeserializeObject<Topo.Model.Progress.GetIntroductionResultsModel>(Awarded)!.results[0].date_awarded, "introduction");
            Assert.AreEqual("2025-06-01", JsonConvert.DeserializeObject<Topo.Model.SIA.GetSIAResultsModel>(Awarded)!.results[0].date_awarded, "SIA");
            Assert.AreEqual("2025-06-01", JsonConvert.DeserializeObject<Topo.Model.AdditionalAwards.GetUnitAchievementsResultModel>(Awarded)!.results[0].date_awarded, "additional award");
            Assert.AreEqual("2025-06-01", JsonConvert.DeserializeObject<Topo.Model.Progress.GetPeakAwardsResultsModel>(Awarded)!.results[0].date_awarded, "peak award");
        }

        [TestMethod]
        public void ARecordWithOnlyAnImportedDateStillGivesThatDate()
        {
            var model = JsonConvert.DeserializeObject<Topo.Model.Milestone.GetMilestoneResultModel>(
                "{\"results\":[{\"status\":\"awarded\",\"status_updated\":\"2026-03-20T09:30:00+11:00\",\"imported\":{\"date_awarded\":\"2024-01-15\"}}]}")!;
            var record = model.results[0];
            Assert.AreEqual(new DateTime(2024, 1, 15), AwardDate.Resolve(record.date_awarded, record.imported?.date_awarded, record.status_updated));
        }
    }
}
