using Topo.Model.Milestone;
using Topo.Model.Wallchart;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class YearsOfService
    {
        private static MilestoneResult Record(params string[] eventDates) => new MilestoneResult
        {
            event_log = eventDates.Select(d => new Event_Log { credit_type = "participant", event_start_datetime = DateTime.Parse(d) }).ToArray()
        };

        [TestMethod]
        public void TheFirstEventIsTheEarliestOneAcrossAllTheSectionsMilestoneRecords()
        {
            var first = YearsOfServiceCalculator.FirstEvent(new[] { Record("2025-03-10T08:00:00", "2024-02-09T08:00:00"), Record("2023-08-01T08:00:00") });
            Assert.AreEqual(new DateTime(2023, 8, 1), first);
        }

        [TestMethod]
        public void TheFirstEventIsADateWithoutATimeOfDay()
        {
            var first = YearsOfServiceCalculator.FirstEvent(new[] { Record("2024-02-09T19:30:00") });
            Assert.AreEqual(new DateTime(2024, 2, 9), first);
        }

        [TestMethod]
        public void LeaderCreditsCountAsWellAsParticipantCredits()
        {
            var record = Record("2025-03-10T08:00:00");
            record.event_log = record.event_log.Append(new Event_Log { credit_type = "leader", event_start_datetime = new DateTime(2024, 5, 4) }).ToArray();
            Assert.AreEqual(new DateTime(2024, 5, 4), YearsOfServiceCalculator.FirstEvent(new[] { record }));
        }

        [TestMethod]
        public void NoLoggedEventGivesNoFirstEvent()
        {
            Assert.IsNull(YearsOfServiceCalculator.FirstEvent(Array.Empty<MilestoneResult>()));
            Assert.IsNull(YearsOfServiceCalculator.FirstEvent(new[] { new MilestoneResult() }));
        }

        [TestMethod]
        public void AnEventWithNoDateIsIgnored()
        {
            var record = Record("2025-03-10T08:00:00");
            record.event_log = record.event_log.Append(new Event_Log { credit_type = "participant" }).ToArray();
            Assert.AreEqual(new DateTime(2025, 3, 10), YearsOfServiceCalculator.FirstEvent(new[] { record }));
        }

        [TestMethod]
        public void YearsRunFromTheFirstEventNotTheIntro()
        {
            var asOf = new DateTime(2026, 11, 22);
            // Intro was a year after they started coming along
            var years = YearsOfServiceCalculator.CalculateCompletedYears(new DateTime(2023, 11, 1), new DateTime(2024, 11, 1), asOf);
            Assert.AreEqual(3, years);
        }

        [TestMethod]
        public void TheIntroIsUsedWhenThereIsNoLoggedEvent()
        {
            var years = YearsOfServiceCalculator.CalculateCompletedYears(null, new DateTime(2024, 11, 1), new DateTime(2026, 11, 22));
            Assert.AreEqual(2, years);
        }

        [TestMethod]
        public void NoFirstEventAndNoIntroGivesNoYears()
        {
            Assert.IsNull(YearsOfServiceCalculator.CalculateCompletedYears(null, null, new DateTime(2026, 11, 22)));
        }

        [TestMethod]
        public void OnlyWholeYearsCount()
        {
            Assert.AreEqual(1, YearsOfServiceCalculator.CalculateCompletedYears(new DateTime(2025, 11, 23), null, new DateTime(2026, 11, 23)));
            Assert.AreEqual(0, YearsOfServiceCalculator.CalculateCompletedYears(new DateTime(2025, 11, 23), null, new DateTime(2026, 11, 22)));
        }
    }
}
