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

        // Terrain dates an award by its UTC day. status_updated is read as a moment in time and shown in the timezone of the
        // device the app runs on (the same way TimeZoneInfo.Local is used for the rest of the app), so these tests work out
        // what the user's own day is from whatever timezone this machine is set to, so they hold in any of them (run them
        // with the TZ environment variable set to Australia/Perth, Australia/Adelaide and so on to try each one).
        private static DateTime AsTheAppReadsIt(string jsonTimestamp) =>
            JsonConvert.DeserializeObject<Topo.Model.Milestone.GetMilestoneResultModel>("{\"results\":[{\"status_updated\":\"" + jsonTimestamp + "\"}]}")!.results[0].status_updated;

        [TestMethod]
        public void TerrainsTimestampsAreReadAsALocalTimeOfTheDeviceTheAppRunsOn()
        {
            var read = AsTheAppReadsIt("2026-08-26T22:25:47.846666+00:00");
            Assert.AreEqual(DateTimeKind.Local, read.Kind);
            Assert.AreEqual(new DateTime(2026, 8, 26, 22, 25, 47, DateTimeKind.Utc).ToLocalTime().Date, read.Date);
        }

        [TestMethod]
        public void AnAwardMadeEarlyInTheMorningGetsTheUsersOwnDayNotTerrainsUtcDay()
        {
            // 22:25 UTC on 26 Aug: that is 8:25am on 27 Aug in Sydney, 6:25am in Perth and 7:55am in Darwin, but still 26 Aug in UTC
            var updated = AsTheAppReadsIt("2026-08-26T22:25:47.846666+00:00");
            var date = AwardDate.Resolve("2026-08-26", "", updated);
            Assert.AreEqual(new DateTime(2026, 8, 26, 22, 25, 47, DateTimeKind.Utc).ToLocalTime().Date, date);
        }

        [TestMethod]
        public void ADateTerrainWroteOnPurposeThatIsNotTheUtcDayOfTheUpdateIsKeptAsItIs()
        {
            // dated the day before it was last touched, in every timezone (24 Jul 05:38 UTC, dated 23 Jul)
            var updated = AsTheAppReadsIt("2025-07-24T05:38:30.873187+00:00");
            Assert.AreEqual(new DateTime(2025, 7, 23), AwardDate.Resolve("2025-07-23", "", updated));
        }

        [TestMethod]
        public void AnImportedDateIsNotMovedToTheUsersDay()
        {
            // imported dates were typed in, not worked out in UTC, so they are never adjusted
            var updated = AsTheAppReadsIt("2026-08-26T22:25:47.846666+00:00");
            Assert.AreEqual(new DateTime(2026, 8, 26), AwardDate.Resolve("", "2026-08-26", updated));
        }

        [TestMethod]
        public void TheResultHasNoTimezoneSoTheReportFunctionRunningInUtcDoesNotMoveItADay()
        {
            var updated = AsTheAppReadsIt("2026-08-26T22:25:47.846666+00:00");
            foreach (var date in new[]
            {
                AwardDate.Resolve("2026-08-26", "", updated),     // the user's day for the award
                AwardDate.Resolve("2025-01-02", "", updated),     // Terrain's date
                AwardDate.Resolve("", "2024-01-15", updated),     // imported
                AwardDate.Resolve("", "", updated)                // when it last changed
            })
            {
                Assert.AreEqual(DateTimeKind.Unspecified, date.Kind);
                var sent = JsonConvert.SerializeObject(date);
                StringAssert.DoesNotMatch(sent, new System.Text.RegularExpressions.Regex(@"[+-]\d\d:\d\d""$|Z""$"));
                Assert.AreEqual(date, JsonConvert.DeserializeObject<DateTime>(sent));
            }
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
