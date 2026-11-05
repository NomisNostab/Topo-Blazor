using Newtonsoft.Json;
using Topo.Model.Members;
using Topo.Model.ReportGeneration;
using TopoReportFunction;
using Topo.Model.Program;
using Topo.Model.UnitReport;
using AdditionalAwardResult = Topo.Model.AdditionalAwards.Result;
using MilestoneResult = Topo.Model.Milestone.MilestoneResult;
using SiaResult = Topo.Model.SIA.Result;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class UnitReport
    {
        private static MemberListModel Member(string id, string number, int adult = 0, string duty = "") =>
            new MemberListModel { id = id, member_number = number, first_name = id, isAdultLeader = adult, patrol_duty = duty };

        private static Attendance_Members Who(string number) => new Attendance_Members { id = "id-" + number, member_number = number };

        private static UnitEventInput Event(string id, string scope, string start, string end, string[] attended, string[] leaders = null!, string status = "confirmed", string[] categories = null!, float walked = 0)
        {
            var evt = new GetEventResultModel
            {
                id = id,
                status = status,
                start_datetime = DateTime.Parse(start),
                end_datetime = DateTime.Parse(end),
                attendance = new Attendance
                {
                    attendee_members = attended.Select(Who).ToArray(),
                    leader_members = (leaders ?? new string[0]).Select(Who).ToArray()
                }
            };
            if (categories != null)
                evt.achievement_pathway_logbook_data = new Achievement_Pathway_Logbook_Data { categories = categories, distance_walkabout = walked };
            return new UnitEventInput { Scope = scope, Event = evt };
        }

        private static UnitReportInput Sample()
        {
            var input = new UnitReportInput
            {
                UnitName = "Scout Unit",
                Section = "scout",
                From = new DateTime(2026, 1, 1),
                To = new DateTime(2026, 6, 30),
                AsOf = new DateTime(2026, 5, 15),
                CurrentMembers = new List<MemberListModel>
                {
                    Member("y1", "1"), Member("y2", "2"), Member("l1", "10", adult: 1), Member("ul", "11", duty: "UL")
                },
                AwardNames = new Dictionary<string, string> { ["aa-1"] = "Bronze Cross" }
            };

            input.Events = new List<UnitEventInput>
            {
                // a weekend camp: two nights, 40 hours, with someone who has since left
                Event("e1", "unit", "2026-03-06T18:00:00", "2026-03-08T10:00:00", new[] { "1", "2", "10", "11", "99" }, new[] { "10" }, categories: new[] { "camping" }),
                Event("e2", "unit", "2026-04-04T09:00:00", "2026-04-04T15:00:00", new[] { "1", "10" }, categories: new[] { "walking_hike" }, walked: 8000),
                // a group event lists other units' members, who are not counted
                Event("e3", "group", "2026-05-01T18:00:00", "2026-05-01T20:00:00", new[] { "1", "55555" }),
                Event("e4", "unit", "2025-12-01T18:00:00", "2025-12-01T20:00:00", new[] { "1" }),
                Event("e5", "unit", "2026-02-01T18:00:00", "2026-02-01T20:00:00", new[] { "1" }, status: "cancelled"),
                Event("e6", "unit", "2026-02-10T18:00:00", "2026-02-10T20:00:00", new string[0]),
                Event("e7", "unit", "2026-06-10T18:00:00", "2026-06-10T20:00:00", new[] { "1" }),
                // the same event twice (two calendars) only counts once
                Event("e2", "unit", "2026-04-04T09:00:00", "2026-04-04T15:00:00", new[] { "1", "10" }, categories: new[] { "walking_hike" }, walked: 8000)
            };

            input.BadgesByMember["y1"] = new UnitMemberBadges
            {
                Milestones = { results = new[]
                {
                    new MilestoneResult { section = "scout", status = "awarded", achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = 1 }, status_updated = new DateTime(2026, 2, 1) },
                    new MilestoneResult { section = "scout", status = "awarded", achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = 2 }, status_updated = new DateTime(2025, 2, 1) }
                } },
                Oas = { results = new[]
                {
                    Oas("scout", "bushcraft", 3, "awarded", new DateTime(2026, 3, 1)),
                    Oas("scout", "camping", 8, "awarded", new DateTime(2026, 3, 1)),
                    Oas("scout", "alpine", 2, "draft_review", new DateTime(2026, 3, 1))
                } },
                Sia = { results = new[] { Sia("scout", "sia_adventure_sports", "awarded", new DateTime(2026, 2, 2)) } },
                Peak = { results = new[] { new Topo.Model.Progress.PeakAwardResult { section = "scout", status = "awarded", status_updated = new DateTime(2026, 4, 1) } } }
            };
            input.BadgesByMember["y2"] = new UnitMemberBadges
            {
                Milestones = { results = new[]
                {
                    // earned as a Cub, so not this unit's
                    new MilestoneResult { section = "cub", status = "awarded", achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = 1 }, status_updated = new DateTime(2026, 2, 1) }
                } },
                Sia = { results = new[] { Sia("scout", "sia_stem_innovation", "awarded", new DateTime(2026, 5, 1)) } }
            };

            input.AdditionalAwards.results = new[]
            {
                Award("y1", "aa-1", new DateTime(2026, 3, 3)),
                Award("y2", "aa-1", new DateTime(2025, 3, 3)),
                Award("someone-else", "aa-1", new DateTime(2026, 3, 3))
            };
            return input;
        }

        private static Topo.Model.OAS.Result Oas(string section, string stream, int stage, string status, DateTime updated) =>
            new Topo.Model.OAS.Result { section = section, status = status, status_updated = updated, achievement_meta = new Topo.Model.OAS.Achievement_Meta { stream = stream, stage = stage } };

        private static SiaResult Sia(string section, string area, string status, DateTime updated) =>
            new SiaResult { section = section, status = status, status_updated = updated, answers = new Topo.Model.SIA.Answers { special_interest_area_selection = area } };

        private static AdditionalAwardResult Award(string member, string awardId, DateTime updated) =>
            new AdditionalAwardResult { member_id = member, section = "scout", status = "awarded", status_updated = updated, achievement_meta = new Topo.Model.AdditionalAwards.Achievement_Meta { additional_award_id = awardId } };

        [TestMethod]
        public void ReportsRenderAsPdfAndExcel()
        {
            var report = UnitReportBuilder.Build(Sample());
            SetEnvironment.SetSyncfusionLicense();
            var function = new Function();
            foreach (var (outputType, extension) in new[] { (OutputType.PDF, "pdf"), (OutputType.Excel, "xlsx") })
            {
                var request = new ReportGenerationRequest
                {
                    ReportType = ReportType.UnitReport,
                    GroupName = "Sample Group",
                    Section = "scout",
                    UnitName = "Scout Unit",
                    OutputType = outputType,
                    ReportData = JsonConvert.SerializeObject(report)
                };
                var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) };
                var result = function.FunctionHandler(proxyRequest, null);
                Assert.IsFalse(string.IsNullOrEmpty(result), $"{extension} came back empty");
                File.WriteAllBytes($@"TestResults/UnitReport.{extension}", Convert.FromBase64String(result));
            }
        }

        [TestMethod]
        public void CountsMembersIncludingFormerAttendeesOfTheUnitsOwnEvents()
        {
            var report = UnitReportBuilder.Build(Sample());

            Assert.AreEqual(2, report.Members.CurrentYouth);
            Assert.AreEqual(2, report.Members.CurrentLeaders); // the unit leader counts as a leader
            Assert.AreEqual(1, report.Members.FormerAttendees); // not the other unit's member from the group event
            Assert.AreEqual(5, report.Members.Total);
        }

        [TestMethod]
        public void OnlyHappenedUncancelledInRangeEventsAreCountedOnce()
        {
            var report = UnitReportBuilder.Build(Sample());

            Assert.AreEqual(4, report.EventsCounted); // e1, e2, e3 and e6
            Assert.AreEqual(1, report.EventsWithoutAttendance);
        }

        [TestMethod]
        public void NightsAndPersonNights()
        {
            var report = UnitReportBuilder.Build(Sample());

            Assert.AreEqual(2, report.NightsRun);
            Assert.AreEqual(10, report.PersonNightsCamped); // five people, two nights
            Assert.AreEqual(6, report.YouthPersonNights);
            Assert.AreEqual(4, report.LeaderPersonNights);
        }

        [TestMethod]
        public void HoursForLeadersAndYouth()
        {
            var report = UnitReportBuilder.Build(Sample());

            // camp 40h: 2 leaders, 3 youth; hike 6h: 1 leader, 1 youth; group event 2h: just the one unit member
            Assert.AreEqual(86.0, report.LeaderHours);
            Assert.AreEqual(128.0, report.YouthHours);
        }

        [TestMethod]
        public void KmHikedOnceOrPerPerson()
        {
            var report = UnitReportBuilder.Build(Sample());

            Assert.AreEqual(1, report.HikeEvents);
            Assert.AreEqual(1, report.HikeEventsWithDistance);
            Assert.AreEqual(8.0, report.HikeKmOnce);
            Assert.AreEqual(16.0, report.HikeKmPerPerson);

            Assert.AreEqual(8.0, report.HikeKm);
            report.KmPerPerson = true;
            Assert.AreEqual(16.0, report.HikeKm);
        }

        [TestMethod]
        public void BadgesAreThisUnitsSectionAndDateRangeOnly()
        {
            var badges = UnitReportBuilder.Build(Sample()).Badges;

            Assert.AreEqual(1, badges.PeakAwards);
            Assert.AreEqual(1, badges.MilestonesTotal); // the 2025 one and the Cub one are left out
            Assert.AreEqual("Milestone 1", badges.Milestones.Single().Name);
            Assert.AreEqual(3, badges.OasProgressions); // bushcraft 1, camping stage 8 counts double; the draft one does not count
            Assert.AreEqual(2, badges.SiaTotal);
            CollectionAssert.AreEqual(new[] { "Adventure Sports", "Stem Innovation" }, badges.SiaCategories.Select(s => s.Name).ToList());
            Assert.AreEqual(1, badges.AdditionalAwardsTotal);
            Assert.AreEqual("Bronze Cross", badges.AdditionalAwards.Single().Name);
        }
    }
}
