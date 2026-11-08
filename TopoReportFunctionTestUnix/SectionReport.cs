using Newtonsoft.Json;
using Topo.Model.History;
using Topo.Model.Members;
using Topo.Model.Program;
using Topo.Model.SectionReport;
using Topo.Model.ReportGeneration;
using Topo.Model.UnitReport;
using TopoReportFunction;
using AdditionalAwardResult = Topo.Model.AdditionalAwards.Result;
using MilestoneResult = Topo.Model.Milestone.MilestoneResult;
using SiaResult = Topo.Model.SIA.Result;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class SectionReport
    {
        private static MemberListModel Member(string id, string number, string first, int adult = 0, string duty = "") =>
            new MemberListModel { id = id, member_number = number, first_name = first, last_name = "X", isAdultLeader = adult, patrol_duty = duty };

        private static Topo.Model.Progress.Result Intro(string section, string awarded) =>
            new Topo.Model.Progress.Result { section = section, status = "awarded", imported = new Topo.Model.Progress.Imported { date_awarded = awarded } };

        // recordSections are the sections of the member's records; a "section@unit" says which unit it was submitted under
        private static SectionMemberRecords Records(string introSection, string introDate, params string[] recordSections)
        {
            var records = new SectionMemberRecords();
            if (introSection != "")
                records.IntroToSection.results = new[] { Intro(introSection, introDate) };
            records.AllAchievements.results = recordSections.Select(s =>
            {
                var parts = s.Split('@');
                return new MemberAchievement
                {
                    section = parts[0],
                    latest_submission = parts.Length > 1 ? new LatestSubmission { unit_id = parts[1] } : null
                };
            }).ToArray();
            return records;
        }

        private static Attendance_Members Who(string number, string first = "") =>
            new Attendance_Members { id = "id-" + number, member_number = number, first_name = first, last_name = "X" };

        private static UnitEventInput Event(string id, string scope, string start, string[] attended, string[] leaders = null!, string status = "confirmed")
        {
            var begin = DateTime.Parse(start);
            return new UnitEventInput
            {
                Scope = scope,
                Event = new GetEventResultModel
                {
                    id = id, title = "Event " + id, status = status, challenge_area = "outdoors",
                    start_datetime = begin, end_datetime = begin.AddHours(3),
                    attendance = new Attendance
                    {
                        attendee_members = attended.Select(n => Who(n, "Visitor" + n)).ToArray(),
                        leader_members = (leaders ?? new string[0]).Select(n => Who(n, "Leader" + n)).ToArray()
                    }
                }
            };
        }

        private static SectionReportInput CubUnit()
        {
            var input = new SectionReportInput
            {
                UnitId = "u-cub", UnitName = "Cub Unit", Section = "cub",
                From = new DateTime(2026, 1, 1), To = new DateTime(2026, 6, 30),
                AsOf = new DateTime(2026, 7, 1),
                CurrentMembers = new List<MemberListModel>
                {
                    Member("y1", "1", "Yara"), Member("y2", "2", "Yuri"), Member("y3", "3", "Yves"), Member("u1", "4", "Una", duty: "UL"),
                    Member("y4", "5", "Zed"), Member("l1", "10", "Lee", adult: 1), Member("l2", "11", "Lou", adult: 1)
                },
                GroupUnitIds = new HashSet<string> { "u-joey", "u-cub", "u-scout" },
                AwardNames = new Dictionary<string, string> { ["aa-1"] = "Service Award" }
            };
            // Yara linked up from the group's Joeys and was invested in the period; Yuri joined off the street in the period;
            // Zed joined in the period from a unit in another group; Yves was invested last year; Una (the unit leader,
            // a youth) is not invested yet
            input.Records["y1"] = Records("cub", "2026-02-10", "joey@u-joey", "cub@u-cub");
            input.Records["y2"] = Records("cub", "2026-03-01", "cub@u-cub");
            input.Records["y3"] = Records("cub", "2025-02-01", "cub@u-cub");
            input.Records["u1"] = Records("", "", "cub@u-cub");
            input.Records["y4"] = Records("cub", "2026-05-01", "cub@u-away");

            // Sam moved up to the Scouts in the period; Pat moved up the year before
            input.NextSectionMembers = new List<NextSectionMember>
            {
                new NextSectionMember { Person = new RosterPerson { Id = "n1", MemberNumber = "20", Name = "Sam Next" }, Records = Records("scout", "2026-04-01", "cub", "scout") },
                new NextSectionMember { Person = new RosterPerson { Id = "n2", MemberNumber = "21", Name = "Pat Next" }, Records = Records("scout", "2025-01-01", "cub", "scout") }
            };

            input.Events = new List<UnitEventInput>
            {
                // a camp with a youth who has since gone (30) and a leader who has since gone (40)
                Event("e1", "unit", "2026-03-06T18:00:00", new[] { "1", "2", "3", "10", "30", "40" }, new[] { "40" }),
                Event("e2", "unit", "2026-04-04T09:00:00", new[] { "1" }, status: "cancelled"),
                Event("e3", "unit", "2026-08-01T18:00:00", new string[0]),
                Event("e4", "unit", "2026-12-01T18:00:00", new string[0]),
                // a group event lists other units' members, who are not this unit's
                Event("e5", "group", "2026-05-01T18:00:00", new[] { "1", "50" })
            };

            input.Previous = new RosterSnapshot
            {
                SavedOn = new DateTime(2026, 1, 5),
                Youth = new List<RosterPerson>
                {
                    new RosterPerson { Id = "y1", MemberNumber = "1", Name = "Yara X" },
                    new RosterPerson { Id = "g1", MemberNumber = "60", Name = "Gone Youth" }
                },
                Leaders = new List<RosterPerson>
                {
                    new RosterPerson { Id = "l1", MemberNumber = "10", Name = "Lee X" },
                    new RosterPerson { Id = "ol", MemberNumber = "70", Name = "Old Leader" }
                }
            };
            return input;
        }

        private static SectionReportSet SetOf(bool group)
        {
            // a fuller unit than most tests need, so the rendered files show every row
            var input = CubUnit();
            AddSampleBadges(input);
            var ages = new[] { "8y 2m", "8y 9m", "9y 1m", "9y 6m", "10y 3m", "10y 10m" };
            input.CurrentMembers.Where(m => m.isAdultLeader == 0).Zip(ages, (m, age) => m.age = age).ToList();
            var helper = Member("h1", "90", "Helper", adult: 1);
            helper.age = "16y 2m";
            input.CurrentMembers.Add(helper);
            var cubs = SectionReportBuilder.Build(input);
            var report = new SectionReportSet
            {
                GeneratedAt = new DateTime(2026, 7, 1, 14, 30, 0), GeneratedBy = "Sample Leader", GroupName = "Sample Group",
                From = new DateTime(2026, 1, 1), To = new DateTime(2026, 6, 30), IsGroup = group,
                Sections = new List<SectionReportModel> { cubs }
            };
            // sorted by the leader on the page
            cubs.Departed.Single(d => d.Name == "Gone Youth").Status = DepartureStatus.MovedGroup;
            if (group)
            {
                var scouts = SectionReportBuilder.Build(CubUnit());
                scouts.UnitName = "Scout Unit";
                scouts.Section = "scout";
                scouts.NextSection = "venturer";
                report.Sections.Add(scouts);
            }
            return report;
        }

        [DataTestMethod]
        [DataRow(false, OutputType.PDF, "pdf")]
        [DataRow(false, OutputType.Excel, "xlsx")]
        [DataRow(true, OutputType.PDF, "pdf")]
        [DataRow(true, OutputType.Excel, "xlsx")]
        public void ReportsRenderAsPdfAndExcel(bool group, OutputType outputType, string extension)
        {
            SetEnvironment.SetSyncfusionLicense();
            var request = new ReportGenerationRequest
            {
                ReportType = ReportType.SectionReport, GroupName = "Sample Group", Section = "cub", UnitName = "Cub Unit",
                OutputType = outputType, ReportData = JsonConvert.SerializeObject(SetOf(group))
            };
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) };

            var result = new Function().FunctionHandler(proxyRequest, null);

            Assert.IsFalse(string.IsNullOrEmpty(result), $"{extension} came back empty");
            Directory.CreateDirectory("TestResults");
            File.WriteAllBytes($@"TestResults/SectionReport_{(group ? "Group" : "Section")}.{extension}", Convert.FromBase64String(result));
        }

        // A unit where no one joined or left, with no events and no badges in the period
        private static SectionReportSet QuietSet()
        {
            var input = new SectionReportInput
            {
                UnitId = "u-cub", UnitName = "Cub Unit", Section = "cub",
                From = new DateTime(2026, 1, 1), To = new DateTime(2026, 6, 30), AsOf = new DateTime(2026, 7, 1),
                CurrentMembers = new List<MemberListModel> { Member("y1", "1", "Yara"), Member("l1", "10", "Lee", adult: 1) },
                GroupUnitIds = new HashSet<string> { "u-cub" }
            };
            input.Records["y1"] = Records("cub", "2025-02-01", "cub@u-cub");
            return new SectionReportSet
            {
                GeneratedAt = new DateTime(2026, 7, 1), GeneratedBy = "Sample Leader", GroupName = "Sample Group",
                From = input.From, To = input.To, Sections = new List<SectionReportModel> { SectionReportBuilder.Build(input) }
            };
        }

        private static List<string> RowLabels(SectionReportSet report)
        {
            SetEnvironment.SetSyncfusionLicense();
            var sheet = new Topo.Services.ReportService().GenerateSectionReportWorkbook(report, "Sample Group", "cub", "Cub Unit").Worksheets[0];
            return Enumerable.Range(1, sheet.UsedRange.LastRow).Select(row => sheet.Range[row, 1].Text).Where(t => !string.IsNullOrEmpty(t)).ToList();
        }

        [TestMethod]
        public void RowsWithNothingInThemAreLeftOut()
        {
            var quiet = QuietSet();
            var unit = quiet.Sections[0];
            Assert.IsFalse(unit.AnyNewMembers);
            Assert.IsFalse(unit.AnyLeaving);
            Assert.IsFalse(unit.AnyEvents);
            Assert.IsFalse(unit.Badges.AnyUsed);

            var labels = RowLabels(quiet);
            foreach (var left in new[] { "New members", "Members leaving", "Youth helpers", "Past events", "Upcoming events", "For new members", "Milestones", "Peak Award" })
                CollectionAssert.DoesNotContain(labels, left);
            foreach (var kept in new[] { "Youth", "Leaders", "Totals" })
                CollectionAssert.Contains(labels, kept);
        }

        [TestMethod]
        public void RowsWithSomethingInThemAreKept()
        {
            var busy = SetOf(false);
            Assert.IsTrue(busy.Sections[0].AnyNewMembers);
            Assert.IsTrue(busy.Sections[0].AnyLeaving);
            Assert.IsTrue(busy.AnyNewMembers);

            var labels = RowLabels(busy);
            foreach (var kept in new[] { "New members", "Members leaving", "Past events", "Milestones" })
                CollectionAssert.Contains(labels, kept);
        }

        [TestMethod]
        public void TheReportSurvivesBeingSentToTheReportFunction()
        {
            var json = JsonConvert.SerializeObject(SetOf(true));
            var back = JsonConvert.DeserializeObject<SectionReportSet>(json)!;

            Assert.AreEqual(2, back.Sections.Count);
            Assert.AreEqual(1, back.Sections[0].MovedGroup);
            Assert.AreEqual(1, back.Sections[0].MovedUp);
            Assert.AreEqual(back.Sections[0].Badges.OasProgressions, SetOf(true).Sections[0].Badges.OasProgressions);
            Assert.AreEqual(DepartureStatus.MovedGroup, back.Sections[0].Departed.Single(d => d.Name == "Gone Youth").Status);
        }

        [TestMethod]
        public void CountsYouthInvestedAndLeaders()
        {
            var report = SectionReportBuilder.Build(CubUnit());

            Assert.AreEqual(5, report.CurrentYouth); // the unit leader is a youth
            Assert.AreEqual(4, report.InvestedYouth);
            Assert.AreEqual(3, report.InvestedInPeriod);
            Assert.AreEqual(2, report.CurrentLeaders);
            Assert.AreEqual("scout", report.NextSection);
        }

        [TestMethod]
        public void NewMembersAreSplitByWhereTheyCameFrom()
        {
            var report = SectionReportBuilder.Build(CubUnit());

            Assert.AreEqual(3, report.NewMembers.Count);
            Assert.AreEqual(1, report.NewFromYoungerSection);
            Assert.AreEqual(1, report.NewOffTheStreet);
            Assert.AreEqual(1, report.NewFromAnotherGroup);
            Assert.AreEqual("Yara X", report.NewMembers.Single(n => n.Origin == NewMemberOrigin.YoungerSection).Name);
            Assert.AreEqual("Yuri X", report.NewMembers.Single(n => n.Origin == NewMemberOrigin.OffTheStreet).Name);
            Assert.AreEqual("Zed X", report.NewMembers.Single(n => n.Origin == NewMemberOrigin.AnotherGroup).Name);
        }

        [TestMethod]
        public void RecordsWithNoUnitOnThemAreTakenToBeFromThisGroup()
        {
            var input = CubUnit();
            input.Records["y1"] = Records("cub", "2026-02-10", "joey", "cub"); // imported records carry no unit
            input.Records["y4"] = Records("cub", "2026-05-01"); // nothing at all

            var report = SectionReportBuilder.Build(input);

            Assert.AreEqual(NewMemberOrigin.YoungerSection, report.NewMembers.Single(n => n.Name == "Yara X").Origin);
            Assert.AreEqual(NewMemberOrigin.OffTheStreet, report.NewMembers.Single(n => n.Name == "Zed X").Origin);
        }

        [TestMethod]
        public void WithoutTheGroupsUnitsAnEarlierSectionRecordMeansFromTheYoungerSection()
        {
            var input = CubUnit();
            input.GroupUnitIds = new HashSet<string>();

            var report = SectionReportBuilder.Build(input);

            // With no list of the group's units, a record under another unit can't be called another group's,
            // so Yara's earlier-section records mean they came up from the younger section
            Assert.AreEqual(0, report.NewFromAnotherGroup);
            Assert.AreEqual(NewMemberOrigin.YoungerSection, report.NewMembers.Single(n => n.Name == "Yara X").Origin);
        }

        [TestMethod]
        public void NewMembersAreGivenTheirInvestitureItems()
        {
            var items = SectionReportBuilder.Build(CubUnit()).Badges.InvestitureItems.ToDictionary(i => i.Name, i => i.Count);

            // one off the street (all six) and one from another group (the first four); the member who linked up gets none
            Assert.AreEqual(2, items["Group Tape"]);
            Assert.AreEqual(2, items["Scarf"]);
            Assert.AreEqual(2, items["Scarf Badge"]);
            Assert.AreEqual(2, items["Region Badge"]);
            Assert.AreEqual(1, items["Australian Scouting Badge"]);
            Assert.AreEqual(1, items["World Scouting Badge"]);
            Assert.AreEqual(6, items.Count);
        }

        [TestMethod]
        public void NoNewMembersMeansNoInvestitureItems()
        {
            var input = CubUnit();
            foreach (var records in input.Records.Values)
                foreach (var intro in records.IntroToSection.results)
                    intro.imported.date_awarded = "2020-01-01";

            Assert.AreEqual(0, SectionReportBuilder.Build(input).Badges.InvestitureItems.Count);
        }

        [TestMethod]
        public void NobodyLinksUpIntoTheYoungestSection()
        {
            // a Joey unit: there is no younger section, so new members are new to Scouting or from another group
            var input = CubUnit();
            input.UnitId = "u-joey";
            input.UnitName = "Joey Unit";
            input.Section = "joey";
            input.Records["y1"] = Records("joey", "2026-02-10", "joey@u-joey");
            input.Records["y2"] = Records("joey", "2026-03-01", "joey@u-joey");
            input.Records["y3"] = Records("joey", "2025-02-01", "joey@u-joey");
            input.Records["u1"] = Records("", "", "joey@u-joey");
            input.Records["y4"] = Records("joey", "2026-05-01", "joey@u-away");

            var report = SectionReportBuilder.Build(input);

            Assert.AreEqual(0, report.NewFromYoungerSection);
            Assert.AreEqual(2, report.NewOffTheStreet);
            Assert.AreEqual(1, report.NewFromAnotherGroup);
            Assert.AreEqual("cub", report.NextSection);
        }

        [TestMethod]
        public void PeopleWhoLeftAreFoundFromTheNextSectionTheSavedRosterAndEvents()
        {
            var report = SectionReportBuilder.Build(CubUnit());

            CollectionAssert.AreEqual(new[] { "Gone Youth", "Sam Next", "Visitor30 X" }, report.Departed.Select(d => d.Name).ToList());
            var sam = report.Departed.Single(d => d.Name == "Sam Next");
            Assert.AreEqual(DepartureStatus.MovedUp, sam.Status); // seen in the next section, so starts as moved up
            Assert.AreEqual("Now in Scouts", sam.Source);
            Assert.AreEqual(DepartureStatus.Unset, report.Departed.Single(d => d.Name == "Gone Youth").Status);
            Assert.AreEqual(DepartureStatus.Unset, report.Departed.Single(d => d.Name == "Visitor30 X").Status);
            Assert.IsFalse(report.Departed.Any(d => d.Name == "Pat Next")); // moved up before the period
            Assert.AreEqual(1, report.MovedUp);
            Assert.AreEqual(2, report.DepartedUnset);
        }

        [TestMethod]
        public void ChoosingWhyTheyLeftChangesTheCounts()
        {
            var report = SectionReportBuilder.Build(CubUnit());
            report.Departed.Single(d => d.Name == "Gone Youth").Status = DepartureStatus.MovedGroup;
            report.Departed.Single(d => d.Name == "Visitor30 X").Status = DepartureStatus.LeftScouting;

            Assert.AreEqual(1, report.MovedUp);
            Assert.AreEqual(1, report.MovedGroup);
            Assert.AreEqual(1, report.LeftScouting);
            Assert.AreEqual(0, report.DepartedUnset);
        }

        [TestMethod]
        public void NewAndDepartedLeadersAreNamedAgainstTheSavedRoster()
        {
            var report = SectionReportBuilder.Build(CubUnit());

            CollectionAssert.AreEqual(new[] { "Lou X" }, report.NewLeaders);
            CollectionAssert.AreEqual(new[] { "Leader40 X", "Old Leader" }, report.DepartedLeaders);
            Assert.AreEqual(new DateTime(2026, 1, 5), report.SavedRosterDate);
        }

        [TestMethod]
        public void WithoutASavedRosterNoLeaderIsCalledNew()
        {
            var input = CubUnit();
            input.Previous = null;

            var report = SectionReportBuilder.Build(input);

            Assert.AreEqual(0, report.NewLeaders.Count);
            Assert.IsNull(report.SavedRosterDate);
            Assert.IsFalse(report.Departed.Any(d => d.Name == "Gone Youth"));
            CollectionAssert.AreEqual(new[] { "Leader40 X" }, report.DepartedLeaders); // still seen at the camp
        }

        [TestMethod]
        public void PastAndUpcomingEventsAreListedWithWhoCame()
        {
            var input = CubUnit();
            input.To = new DateTime(2026, 9, 30);
            var report = SectionReportBuilder.Build(input);

            // the cancelled one and the group event's other units' members aside
            CollectionAssert.AreEqual(new[] { "Event e1", "Event e5" }, report.PastEvents.Select(e => e.Title).ToList());
            var camp = report.PastEvents.First();
            Assert.AreEqual(4, camp.YouthAttended); // Yara, Yuri, Yves and the youth who has gone
            Assert.AreEqual(2, camp.LeadersAttended); // Lee and the leader who has gone
            Assert.AreEqual(1, report.PastEvents.Last().YouthAttended); // only this unit's own member at the group event
            CollectionAssert.AreEqual(new[] { "Event e3" }, report.UpcomingEvents.Select(e => e.Title).ToList());
        }

        [TestMethod]
        public void UpcomingEventsStopAtTheEndOfThePeriod()
        {
            var input = CubUnit();
            input.To = new DateTime(2026, 7, 15); // before the event on 1 August

            Assert.AreEqual(0, SectionReportBuilder.Build(input).UpcomingEvents.Count);
        }

        [TestMethod]
        public void BadgesUsedInThePeriodForThisSectionOnly()
        {
            var input = CubUnit();
            AddSampleBadges(input);

            var badges = SectionReportBuilder.Build(input).Badges;

            Assert.AreEqual(1, badges.PeakAwards);
            Assert.AreEqual(1, badges.MilestonesTotal);
            Assert.AreEqual(1, badges.SiaTotal);
            Assert.AreEqual("Environment", badges.SiaCategories.Single().Name);
            Assert.AreEqual(1, badges.AdditionalAwardsTotal);
            Assert.AreEqual("Service Award", badges.AdditionalAwards.Single().Name);
            // Yara was awarded bushcraft stages 1 and 2 on the same day, so only needs the stage 2 badge; Yuri has bushcraft
            // stage 1 and camping stage 7 (and an approved alpine stage 3 that has not been handed over)
            Assert.AreEqual(1, badges.OasStages.Single(o => o.Stream == "Bushcraft" && o.Stage == 2).Count);
            Assert.AreEqual(1, badges.OasStages.Single(o => o.Stream == "Bushcraft" && o.Stage == 1).Count);
            Assert.AreEqual(1, badges.OasStages.Single(o => o.Stream == "Camping" && o.Stage == 7).Count);
            Assert.AreEqual(3, badges.OasBadgesTotal);
            // progressions count every stage awarded and stages 7 and above twice: 1 + 1 (Yara), 1 + 2 (Yuri)
            Assert.AreEqual(5, badges.OasProgressions);
        }

        // Badges awarded in the period to Yara and Yuri, one approved but not handed over, and some from other sections
        private static void AddSampleBadges(SectionReportInput input)
        {
            input.Records["y1"].Badges = new Topo.Model.UnitReport.UnitMemberBadges
            {
                Milestones = { results = new[]
                {
                    new MilestoneResult { section = "cub", status = "awarded", achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = 1 }, status_updated = new DateTime(2026, 3, 1) },
                    new MilestoneResult { section = "joey", status = "awarded", achievement_meta = new Topo.Model.Milestone.Achievement_Meta { stage = 1 }, status_updated = new DateTime(2026, 3, 1) }
                } },
                Sia = { results = new[] { new SiaResult { section = "cub", status = "awarded", status_updated = new DateTime(2026, 2, 2), answers = new Topo.Model.SIA.Answers { special_interest_area_selection = "sia_environment" } } } },
                Peak = { results = new[] { new Topo.Model.Progress.PeakAwardResult { section = "cub", status = "awarded", status_updated = new DateTime(2026, 4, 1) } } }
            };
            input.Records["y1"].Badges.Oas.results = new[] { Oas("bushcraft", 1), Oas("bushcraft", 2) };
            var approved = Oas("alpine", 3);
            approved.status = "approved"; // approved but not handed over yet, so not counted
            input.Records["y2"].Badges.Oas.results = new[] { Oas("bushcraft", 1), Oas("camping", 7), approved };
            input.AdditionalAwards.results = new[]
            {
                new AdditionalAwardResult { member_id = "y2", section = "cub", status = "awarded", status_updated = new DateTime(2026, 5, 5), achievement_meta = new Topo.Model.AdditionalAwards.Achievement_Meta { additional_award_id = "aa-1" } },
                new AdditionalAwardResult { member_id = "y2", section = "cub", status = "awarded", status_updated = new DateTime(2025, 5, 5), achievement_meta = new Topo.Model.AdditionalAwards.Achievement_Meta { additional_award_id = "aa-1" } }
            };
        }

        private static Topo.Model.OAS.Result Oas(string stream, int stage, int month = 3) =>
            new Topo.Model.OAS.Result { section = "cub", status = "awarded", status_updated = new DateTime(2026, month, 1), achievement_meta = new Topo.Model.OAS.Achievement_Meta { stream = stream, stage = stage } };

        [TestMethod]
        public void JumpingSeveralOasLevelsAtOnceNeedsOnlyTheHighestBadge()
        {
            var input = CubUnit();
            // stages 2, 3 and 4 all awarded on the same day, then stage 5 a month later, and stage 1 of another stream
            input.Records["y1"].Badges.Oas.results = new[] { Oas("camping", 2), Oas("camping", 3), Oas("camping", 4), Oas("camping", 5, month: 4), Oas("alpine", 1) };

            var stages = SectionReportBuilder.Build(input).Badges.OasStages;

            Assert.AreEqual(3, stages.Sum(s => s.Count));
            Assert.AreEqual(1, stages.Single(s => s.Stream == "Camping" && s.Stage == 4).Count);
            Assert.AreEqual(1, stages.Single(s => s.Stream == "Camping" && s.Stage == 5).Count);
            Assert.AreEqual(1, stages.Single(s => s.Stream == "Alpine" && s.Stage == 1).Count);
            Assert.IsFalse(stages.Any(s => s.Stream == "Camping" && (s.Stage == 2 || s.Stage == 3)));
        }

        [TestMethod]
        public void ASnapshotRemembersYouthAndLeadersSeparately()
        {
            var snapshot = SectionReportBuilder.TakeSnapshot(CubUnit().CurrentMembers, new DateTime(2026, 7, 1));

            Assert.AreEqual(5, snapshot.Youth.Count);
            Assert.AreEqual(2, snapshot.Leaders.Count);
            Assert.IsTrue(snapshot.Youth.Any(p => p.Name == "Una X"));
        }

        [TestMethod]
        public void YouthHelpersAreNeitherLeadersNorYouth()
        {
            var input = CubUnit();
            var before = SectionReportBuilder.Build(input);
            var helper = Member("h1", "90", "Helper", adult: 1);
            helper.age = "16y 2m";
            var adult = Member("h2", "91", "Adult", adult: 1);
            adult.age = "49y 5m";
            input.CurrentMembers.Add(helper);
            input.CurrentMembers.Add(adult);

            var report = SectionReportBuilder.Build(input);

            Assert.AreEqual(before.CurrentLeaders + 1, report.CurrentLeaders);
            Assert.AreEqual(before.CurrentYouth, report.CurrentYouth);
            Assert.IsFalse(report.Leaders.Any(l => l.Name == "Helper X"));
            Assert.IsTrue(report.Leaders.Any(l => l.Name == "Adult X"));
            CollectionAssert.AreEqual(new[] { "Helper X" }, report.YouthHelpers.Select(h => h.Name).ToList());
            Assert.IsFalse(SectionReportBuilder.TakeSnapshot(input.CurrentMembers, DateTime.Today).Leaders.Any(p => p.Name == "Helper X"));
        }

        [TestMethod]
        public void OnlyTheAdultsInTheLeadersPatrolAreTheUnitsLeaders()
        {
            var input = CubUnit();
            var before = SectionReportBuilder.Build(input);
            foreach (var leader in input.CurrentMembers.Where(m => m.isAdultLeader == 1))
                leader.patrol_name = "Adult Leaders - Cubs";
            var advisor = Member("g1", "92", "Advisor", adult: 1); // a group role: the leader duty but no patrol
            advisor.age = "70y 0m";
            input.CurrentMembers.Add(advisor);

            var report = SectionReportBuilder.Build(input);

            Assert.AreEqual(before.CurrentLeaders, report.CurrentLeaders);
            Assert.IsFalse(report.Leaders.Any(l => l.Name == "Advisor X"));
            Assert.AreEqual(before.CurrentYouth, report.CurrentYouth);
        }

        private static List<MemberListModel> YouthAged(params string[] ages) =>
            ages.Select((age, i) => { var m = Member($"a{i}", $"{100 + i}", $"Kid{i}"); m.age = age; return m; }).ToList();

        [TestMethod]
        public void AgesGiveTheAverageAndHowEvenlyTheyAreSpread()
        {
            var even = SectionReportBuilder.AgeProfileOf(YouthAged("11y 6m", "11y 0m", "12y 6m", "12y 0m", "13y 6m", "13y 0m", "14y 6m", "14y 0m"), "scout");

            Assert.AreEqual(8, even.Counted);
            CollectionAssert.AreEqual(new[] { 11, 12, 13, 14 }, even.Buckets.Select(b => b.Age).ToList());
            CollectionAssert.AreEqual(new[] { 2, 2, 2, 2 }, even.Buckets.Select(b => b.Count).ToList());
            Assert.AreEqual(1.0, even.Evenness!.Value, 0.001);
            Assert.AreEqual("12y 9m", AgeProfile.Format(even.AverageYears));
            StringAssert.StartsWith(even.Spread, "Evenly spread");

            var bunched = SectionReportBuilder.AgeProfileOf(YouthAged("11y 1m", "11y 2m", "11y 3m", "11y 4m", "11y 5m", "14y 0m"), "scout");
            Assert.IsTrue(bunched.Evenness < 0.65);
            StringAssert.Contains(bunched.Spread, "Bunched together");
            StringAssert.Contains(bunched.Spread, "leaning younger");
        }

        [TestMethod]
        public void AgesAreWidenedForAnyoneOutsideTheSectionAndSafeWithFewOrNone()
        {
            var wide = SectionReportBuilder.AgeProfileOf(YouthAged("10y 11m", "12y 0m", "15y 3m"), "scout");
            CollectionAssert.AreEqual(new[] { 10, 11, 12, 13, 14, 15 }, wide.Buckets.Select(b => b.Age).ToList());

            var few = SectionReportBuilder.AgeProfileOf(YouthAged("12y 0m", "13y 0m"), "scout");
            Assert.IsNull(few.Evenness);
            Assert.AreEqual("Too few members to say", few.Spread);

            var none = SectionReportBuilder.AgeProfileOf(new List<MemberListModel>(), "scout");
            Assert.AreEqual(0, none.Counted);
            Assert.AreEqual(0, none.Buckets.Count);
            Assert.AreEqual(0, SectionReportBuilder.AgeProfileOf(YouthAged("", "unknown"), "cub").Counted);
        }

        [TestMethod]
        public void AMemberListedTwiceIsOnePerson()
        {
            var before = SectionReportBuilder.Build(CubUnit());
            var input = CubUnit();
            input.CurrentMembers.Add(Member("y1", "1", "Yara"));
            input.CurrentMembers.Add(Member("l1", "10", "Lee", adult: 1));

            var report = SectionReportBuilder.Build(input);

            Assert.AreEqual(before.CurrentYouth, report.CurrentYouth);
            Assert.AreEqual(before.CurrentLeaders, report.CurrentLeaders);
            Assert.AreEqual(before.PastEvents.Count, report.PastEvents.Count);
            var snapshot = SectionReportBuilder.TakeSnapshot(input.CurrentMembers, DateTime.Today);
            Assert.AreEqual(5, snapshot.Youth.Count);
            Assert.AreEqual(2, snapshot.Leaders.Count);
        }

        [TestMethod]
        public void UnitsWithNoYouthNoLeadersOrNothingAtAllStillBuild()
        {
            var adultsOnly = CubUnit();
            adultsOnly.CurrentMembers = adultsOnly.CurrentMembers.Where(m => m.isAdultLeader == 1).ToList();
            var report = SectionReportBuilder.Build(adultsOnly);
            Assert.AreEqual(0, report.CurrentYouth);
            Assert.AreEqual(0, report.InvestedYouth);
            Assert.AreEqual(2, report.CurrentLeaders);

            var noAdults = CubUnit();
            noAdults.CurrentMembers = noAdults.CurrentMembers.Where(m => m.isAdultLeader == 0).ToList();
            Assert.AreEqual(0, SectionReportBuilder.Build(noAdults).CurrentLeaders);

            var nothing = new SectionReportInput { Section = "scout", From = new DateTime(2026, 1, 1), To = new DateTime(2026, 6, 30) };
            var empty = SectionReportBuilder.Build(nothing);
            Assert.AreEqual(0, empty.CurrentYouth + empty.CurrentLeaders + empty.PastEvents.Count + empty.Departed.Count);
        }

        [TestMethod]
        public void TheLastSectionHasNoNextSectionToMoveUpTo()
        {
            Assert.AreEqual("cub", SectionReportBuilder.NextSectionOf("Joey"));
            Assert.AreEqual("", SectionReportBuilder.NextSectionOf("rover"));
            Assert.AreEqual("", SectionReportBuilder.NextSectionOf("unknown"));
        }

        [TestMethod]
        public void GroupTotalsAddUpEachUnit()
        {
            var cubs = SectionReportBuilder.Build(CubUnit());
            var scouts = SectionReportBuilder.Build(CubUnit());
            scouts.Departed.ForEach(d => d.Status = DepartureStatus.LeftScouting);
            var set = new SectionReportSet { IsGroup = true, Sections = new List<SectionReportModel> { cubs, scouts } };

            Assert.AreEqual(10, set.TotalYouth);
            Assert.AreEqual(2, set.TotalLeaders); // the same two adults lead both units
            Assert.AreEqual(6, set.TotalInvestedInPeriod);
            Assert.AreEqual(2, set.TotalNewFromYounger);
            Assert.AreEqual(2, set.TotalNewFromAnotherGroup);
            Assert.AreEqual(4, set.TotalInvestitureItems.Single(i => i.Name == "Scarf").Count);
            Assert.AreEqual(1, set.TotalMovedUp);
            Assert.AreEqual(3, set.TotalLeftScouting);

            scouts.Leaders[0].MemberNumber = "different";
            Assert.AreEqual(3, set.TotalLeaders);
        }
    }
}
