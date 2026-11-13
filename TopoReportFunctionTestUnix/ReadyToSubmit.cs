using Newtonsoft.Json;
using Syncfusion.Pdf.Parsing;
using Topo.Model.Members;
using Topo.Model.ReadyToSubmit;
using Topo.Model.ReportGeneration;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class ReadyToSubmit
    {
        private const string PlanTemplate = @"{""document"":[
            {""id"":""project_name"",""type"":""text_input"",""label"":""Name your project*""},
            {""id"":""Goal_vision"",""type"":""text_input"",""label"":""My goal is...""},
            {""id"":""which_goals"",""type"":""multi_tab_selector"",""label"":""If so, which one(s)?""},
            {""id"":""intro"",""type"":""markdown_text"",""label"":""Read this""},
            {""id"":""attach"",""type"":""file_uploader"",""label"":""Files""},
            {""id"":""agree"",""type"":""checkbox"",""label"":""I agree""}]}";

        private const string ReviewTemplate = @"{""document"":[{""id"":""highlight_SIA"",""type"":""text_input"",""label"":""Highlight""}]}";

        private const string OasTemplateJson = @"{""document"":[{""title"":""Stage 3"",""input_groups"":[
            {""title"":""Plan>"",""inputs"":[{""id"":""plan_one"",""label"":""One""},{""id"":""file_uploader"",""label"":""Files""}]},
            {""title"":""Do>"",""inputs"":[{""id"":""do_one"",""label"":""Two""},{""id"":""logbook_up_to_date"",""label"":""Logbook""}]}]}]}";

        // A stage with four things to tick off, the last the Scout's own logbook tick
        private const string FourInputTemplate = @"{""document"":[{""title"":""Stage 2"",""input_groups"":[{""title"":""Do>"",""inputs"":[
            {""id"":""a"",""label"":""Item A""},{""id"":""b"",""label"":""Item B""},{""id"":""c"",""label"":""Item C""},{""id"":""logbook_up_to_date"",""label"":""My logbook is up to date""}]}]}]}";

        private static ReadyTemplates Templates() => new ReadyTemplates
        {
            Forms =
            {
                ["sia/cub/plan"] = ReadyParsers.ParseFormQuestions(PlanTemplate),
                ["sia/cub/review"] = ReadyParsers.ParseFormQuestions(ReviewTemplate),
                ["personal-development-course/cub"] = ReadyParsers.ParseFormQuestions(@"{""document"":[{""id"":""what_did_you_learn"",""type"":""text_input"",""label"":""Learn""}]}")
            },
            OasInputs =
            {
                ["oas/bushcraft/3"] = ReadyParsers.ParseOasInputs(OasTemplateJson),
                ["oas/camping/2"] = ReadyParsers.ParseOasInputs(FourInputTemplate)
            }
        };

        private static AchievementRecord Record(string type, string status, string section = "cub", string template = "", params (string key, string value)[] answers)
        {
            var record = new AchievementRecord { Id = Guid.NewGuid().ToString(), Type = type, Status = status, Section = section, Template = template, StatusUpdated = new DateTime(2026, 10, 3) };
            foreach (var (key, value) in answers)
                record.Answers[key] = value;
            return record;
        }

        private static AchievementRecord Milestone(string status, string requirements, int stage = 1, string section = "cub")
        {
            var record = Record("milestone", status, section);
            record.RequirementStatus = requirements;
            record.Meta["stage"] = stage.ToString();
            return record;
        }

        private static AchievementRecord Oas(string status, string stream, int stage, params (string key, string value)[] answers)
        {
            var record = Record("outdoor_adventure_skill", status, "cub", "oas/bushcraft/3", answers);
            record.Meta["stream"] = stream;
            record.Meta["stage"] = stage.ToString();
            return record;
        }

        private static ReadyMemberInput Member(string name, params AchievementRecord[] records) =>
            new ReadyMemberInput { MemberId = name, MemberName = name, Records = records.ToList() };

        private static ReadyToSubmitReport Build(params ReadyMemberInput[] members) =>
            ReadyToSubmitBuilder.Build(new ReadyToSubmitInput { UnitName = "Cub Unit", Section = "cub", Members = members.ToList(), Templates = Templates() });

        [TestMethod]
        public void TerrainsAchievementsAreReadWithTheirAnswersAndMeta()
        {
            var json = @"{""results"":[{""id"":""a1"",""type"":""outdoor_adventure_skill"",""section"":""Cub"",""template"":""oas/bushcraft/3"",""status"":""draft_review"",
                ""status_updated"":""2026-10-03T01:02:03.000Z"",""answers"":{""plan_one_verifiedDate"":""2026-09-01"",""logbook_up_to_date"":true,""list"":[""a"",""b""],""nested"":{""x"":1}},
                ""achievement_meta"":{""stream"":""bushcraft"",""stage"":3}}]}";

            var record = ReadyParsers.ParseAchievements(json).Single();

            Assert.AreEqual("cub", record.Section);
            Assert.AreEqual("draft_review", record.Status);
            Assert.AreEqual("2026-09-01", record.Answers["plan_one_verifiedDate"]);
            Assert.AreEqual("true", record.Answers["logbook_up_to_date"]);
            Assert.AreEqual("a, b", record.Answers["list"]);
            Assert.AreEqual(@"{""x"":1}", record.Answers["nested"]); // a date range and the like count as an answer
            Assert.AreEqual("3", record.Meta["stage"]);
            Assert.AreEqual(2026, record.StatusUpdated.Year);
            Assert.AreEqual(0, ReadyParsers.ParseAchievements("not json").Count);
            Assert.AreEqual(0, ReadyParsers.ParseAchievements("{}").Count);
        }

        [TestMethod]
        public void ARecordOfTheWholeUnitCanBeToldApartByWhoseItIs()
        {
            var json = @"{""results"":[
                {""id"":""r1"",""member_id"":""m1"",""type"":""milestone"",""section"":""cub"",""status"":""in_progress"",""status_updated"":""2026-10-03T01:02:03.000Z""},
                {""id"":""r2"",""member_id"":""m2"",""type"":""milestone"",""section"":""cub"",""status"":""in_progress"",""status_updated"":""2026-10-03T01:02:03.000Z""},
                {""id"":""r3"",""member_id"":""m1"",""type"":""milestone"",""section"":""cub"",""status"":""awarded"",""status_updated"":""2026-10-03T01:02:03.000Z""}]}";

            var byMember = ReadyParsers.ParseAchievements(json).ToLookup(r => r.MemberId);

            CollectionAssert.AreEqual(new[] { "r1", "r3" }, byMember["m1"].Select(r => r.Id).ToList());
            CollectionAssert.AreEqual(new[] { "r2" }, byMember["m2"].Select(r => r.Id).ToList());
            Assert.AreEqual(0, byMember["someone not in the unit"].Count());
        }

        [TestMethod]
        public void AMilestoneWithItsRequirementsCompleteButNotSentIsReady()
        {
            var report = Build(Member("Ann",
                Milestone("in_progress", "complete", 1),
                Milestone("in_progress", "incomplete", 2),
                Milestone("awarded", "complete", 3),
                Milestone("pending_review", "complete", 3),
                Milestone("in_progress", "complete", 2, section: "joey")));

            Assert.AreEqual("Milestone 1", report.Items.Single().Badge);
            Assert.AreEqual(ReadyCategory.Milestone, report.Items.Single().Category);
        }

        [TestMethod]
        public void AnOasStageIsReadyWhenEveryRequirementIsTickedOffButItIsStillADraft()
        {
            var done = new[] { ("plan_one_verifiedDate", "2026-09-01"), ("do_one_verifiedDate", "2026-09-02"), ("logbook_up_to_date", "true") };
            var report = Build(
                Member("Ann", Oas("draft_review", "bushcraft", 3, done)),
                Member("Ben", Oas("draft_review", "bushcraft", 3, ("plan_one_verifiedDate", "2026-09-01"))), // not finished
                Member("Cat", Oas("pending_review", "bushcraft", 3, done)), // already sent
                Member("Dan", Oas("draft_review", "bushcraft", 3, ("plan_one_verifiedDate", "2026-09-01"), ("do_one_verifiedDate", "2026-09-02"), ("logbook_up_to_date", "false"))));

            Assert.AreEqual("Ann", report.Items.Single().MemberName);
            Assert.AreEqual("Bushcraft stage 3", report.Items.Single().Badge);
        }

        [TestMethod]
        public void AStageOnlyWaitingOnTheScoutsLogbookTickIsListedOnlyWhenAsked()
        {
            var member = Member("Ann", Oas("draft_review", "bushcraft", 3, ("plan_one_verifiedDate", "2026-09-01"), ("do_one_verifiedDate", "2026-09-02")));
            ReadyToSubmitReport Run(bool optional) => ReadyToSubmitBuilder.Build(new ReadyToSubmitInput
            {
                Section = "cub", Members = new List<ReadyMemberInput> { member }, Templates = Templates(), LogbookTickOptional = optional
            });

            Assert.AreEqual(0, Run(false).Items.Count);
            var item = Run(true).Items.Single();
            StringAssert.Contains(item.Waiting, "logbook");

            // two things left is not just the logbook
            var further = Member("Ben", Oas("draft_review", "bushcraft", 3, ("plan_one_verifiedDate", "2026-09-01")));
            Assert.AreEqual(0, ReadyToSubmitBuilder.Build(new ReadyToSubmitInput { Section = "cub", Members = new List<ReadyMemberInput> { further }, Templates = Templates(), LogbookTickOptional = true }).Items.Count);
        }

        // A camping stage 2 draft with these requirements ticked off
        private static AchievementRecord Camping(params string[] ticked)
        {
            var record = Oas("draft_review", "camping", 2, ticked.Select(id => id == "logbook_up_to_date" ? (id, "true") : (id + "_verifiedDate", "2026-09-01")).ToArray());
            record.Template = "oas/camping/2";
            return record;
        }

        [TestMethod]
        public void StagesAFewTicksShortAreListedCommonestMissingItemFirst()
        {
            var members = new[]
            {
                Member("Ann", Camping("a", "b", "logbook_up_to_date")),           // missing C
                Member("Ben", Camping("a", "b", "logbook_up_to_date")),           // missing C
                Member("Cat", Camping("b", "c", "logbook_up_to_date")),           // missing A
                Member("Dan", Camping("c", "logbook_up_to_date")),                // missing A and B
                Member("Eve", Camping("a", "b", "c", "logbook_up_to_date")),      // done: ready
                Member("Fay", Camping("a")),                                      // missing three
            };
            ReadyToSubmitReport Run(int max) => ReadyToSubmitBuilder.Build(new ReadyToSubmitInput
            {
                Section = "cub", Members = members.ToList(), Templates = Templates(), MaxTicksShort = max
            });

            Assert.AreEqual("Eve", Run(2).Ready.Single().MemberName);
            Assert.AreEqual(0, Run(0).NearlyThere.Count);
            CollectionAssert.AreEqual(new[] { "Ann", "Ben", "Cat" }, Run(1).NearlyThere.Select(i => i.MemberName).ToList());
            // C is missing for two Scouts and A for two (Cat and Dan): ties keep the template's order, so Item A first, and Dan, missing A
            // as well as B, goes with A
            var two = Run(2).NearlyThere;
            CollectionAssert.AreEqual(new[] { "Cat", "Dan", "Ann", "Ben" }, two.Select(i => i.MemberName).ToList());
            CollectionAssert.AreEqual(new[] { "Item A" }, two[0].Missing);
            CollectionAssert.AreEqual(new[] { "Item A", "Item B" }, two[1].Missing);
            Assert.AreEqual(2, two[1].TicksShort);
            CollectionAssert.AreEqual(new[] { "Item C" }, two[2].Missing);
            // three short is more than asked for
            Assert.IsFalse(Run(2).Items.Any(i => i.MemberName == "Fay"));
            Assert.IsTrue(Run(3).Items.Any(i => i.MemberName == "Fay"));
        }

        [TestMethod]
        public void ThePlainlyCommonestMissingItemComesBeforeLessCommonOnes()
        {
            var members = new[]
            {
                Member("Ann", Camping("a", "c", "logbook_up_to_date")),   // missing B
                Member("Ben", Camping("a", "b", "logbook_up_to_date")),   // missing C
                Member("Cat", Camping("a", "c", "logbook_up_to_date")),   // missing B
                Member("Dan", Camping("a", "c", "logbook_up_to_date")),   // missing B
            };
            var near = ReadyToSubmitBuilder.Build(new ReadyToSubmitInput { Section = "cub", Members = members.ToList(), Templates = Templates(), MaxTicksShort = 1 }).NearlyThere;

            CollectionAssert.AreEqual(new[] { "Ann", "Cat", "Dan", "Ben" }, near.Select(i => i.MemberName).ToList()); // three miss B, one misses C
        }

        [TestMethod]
        public void NearlyThereStagesAreSplitByBadgeWithAGridOfWhatEachScoutIsMissing()
        {
            var vertical = Oas("draft_review", "vertical", 2, ("plan_one_verifiedDate", "2026-09-01"));
            vertical.Template = "oas/bushcraft/3"; // the same plan/do template, standing in for another stage
            var members = new[]
            {
                Member("Ann", Camping("a", "b", "logbook_up_to_date")),                 // camping: missing C
                Member("Ben", Camping("b", "c", "logbook_up_to_date")),                 // camping: missing A
                Member("Cat", Camping("c", "logbook_up_to_date")),                      // camping: missing A and B
                Member("Dan", vertical),                                                // vertical stage 2: missing three
                Member("Eve", Camping("a", "b", "c", "logbook_up_to_date"))             // done
            };
            var report = ReadyToSubmitBuilder.Build(new ReadyToSubmitInput { Section = "cub", Members = members.ToList(), Templates = Templates(), MaxTicksShort = 3 });

            var badges = report.NearlyThereByBadge();

            // in the wall chart's order, and each badge's Scouts stay in the unit's order
            CollectionAssert.AreEqual(new[] { "Camping stage 2", "Vertical stage 2" }, badges.Select(b => b.Badge).ToList());
            var camping = badges[0];
            CollectionAssert.AreEqual(new[] { "Ann", "Ben", "Cat" }, camping.Scouts.Select(s => s.MemberName).ToList());
            // Item A is missing for two Scouts and B for one, C for one: A first, then the ties in the template's order
            CollectionAssert.AreEqual(new[] { "Item A", "Item B", "Item C" }, camping.Requirements.Select(r => r.Label).ToList());
            CollectionAssert.AreEqual(new[] { 2, 1, 1 }, camping.Requirements.Select(r => r.Scouts).ToList());
            var itemA = camping.Requirements[0];
            Assert.IsFalse(camping.IsMissing(camping.Scouts[0], itemA)); // Ann has A: blocked out
            Assert.IsTrue(camping.IsMissing(camping.Scouts[1], itemA));  // Ben still has to tick it
            Assert.IsTrue(camping.IsMissing(camping.Scouts[2], itemA));
            Assert.AreEqual(0, new ReadyToSubmitReport().NearlyThereByBadge().Count);
        }

        [TestMethod]
        public void NearlyThereBadgesFollowTheWallChartThenLowestLevelFirst()
        {
            AchievementRecord Stage(string stream, int stage, string template)
            {
                var record = Oas("draft_review", stream, stage, ("a_verifiedDate", "2026-09-01"));
                record.Template = template;
                return record;
            }
            // two Scouts on vertical, one each on the others, and the unit's order is not the wall chart's
            var members = new[]
            {
                Member("Ann", Stage("vertical", 1, "oas/camping/2"), Stage("camping", 2, "oas/camping/2")),
                Member("Ben", Stage("vertical", 1, "oas/camping/2"), Stage("paddling", 3, "oas/camping/2")),
                Member("Cat", Stage("camping", 1, "oas/camping/2"), Stage("bushcraft", 4, "oas/camping/2"), Stage("aquatics", 2, "oas/camping/2"))
            };
            var report = ReadyToSubmitBuilder.Build(new ReadyToSubmitInput { Section = "cub", Members = members.ToList(), Templates = Templates(), MaxTicksShort = 3 });

            CollectionAssert.AreEqual(
                new[] { "Bushcraft stage 4", "Camping stage 1", "Camping stage 2", "Vertical stage 1", "Aquatics stage 2", "Paddling stage 3" },
                report.NearlyThereByBadge().Select(b => b.Badge).ToList());
        }

        [TestMethod]
        public void ScoutsAreNamedByFirstNameWithEnoughSurnameToTellThemApartUnlessFullNamesAreAsked()
        {
            ReadyMemberInput Person(string first, string last) => new ReadyMemberInput { MemberId = first + last, FirstName = first, LastName = last, MemberName = first + " " + last };
            var members = new[] { Person("Sam", "Smith"), Person("Sam", "Stone"), Person("Sam", "Brown"), Person("Ava", "Lee"), Person("Zoe", "Hall"), Person("Zoe", "Hall") };

            var short_ = ReadyToSubmitBuilder.DisplayNames(members, false);
            CollectionAssert.AreEqual(
                new[] { "Sam Sm", "Sam St", "Sam Br", "Ava", "Zoe Hall (1)", "Zoe Hall (2)" },
                members.Select(m => short_[m]).ToList());

            var full = ReadyToSubmitBuilder.DisplayNames(members, true);
            CollectionAssert.AreEqual(
                new[] { "Sam Smith", "Sam Stone", "Sam Brown", "Ava Lee", "Zoe Hall (1)", "Zoe Hall (2)" },
                members.Select(m => full[m]).ToList());

            // the report names Scouts the short way unless asked, and keeps a name given as one piece
            var oneName = new ReadyMemberInput { MemberId = "x", MemberName = "Ann Lee", Records = new List<AchievementRecord> { Milestone("in_progress", "complete") } };
            Assert.AreEqual("Ann Lee", Build(oneName).Items.Single().MemberName);
            var named = new ReadyMemberInput { MemberId = "y", FirstName = "Ann", LastName = "Lee", MemberName = "Ann Lee", Records = new List<AchievementRecord> { Milestone("in_progress", "complete") } };
            Assert.AreEqual("Ann", Build(named).Items.Single().MemberName);
        }

        [TestMethod]
        public void TheLogbookTickIsItsOwnListAndNotAlsoNearlyThere()
        {
            var members = new[] { Member("Ann", Camping("a", "b", "c")), Member("Ben", Camping("a", "b", "logbook_up_to_date")) };
            ReadyToSubmitReport Run(bool logbook, int max) => ReadyToSubmitBuilder.Build(new ReadyToSubmitInput
            {
                Section = "cub", Members = members.ToList(), Templates = Templates(), LogbookTickOptional = logbook, MaxTicksShort = max
            });

            var both = Run(true, 1);
            Assert.AreEqual("Ann", both.LogbookOnly.Single().MemberName);
            Assert.AreEqual("Ben", both.NearlyThere.Single().MemberName);
            // with the logbook option off, a stage short only that tick is simply one tick short
            var near = Run(false, 1).NearlyThere;
            CollectionAssert.AreEqual(new[] { "Ann", "Ben" }.OrderBy(n => n).ToList(), near.Select(i => i.MemberName).OrderBy(n => n).ToList());
            Assert.AreEqual(0, Run(false, 1).LogbookOnly.Count);
            StringAssert.Contains(near.First(i => i.MemberName == "Ann").Missing.Single(), "logbook");
        }

        [TestMethod]
        public void ADateRangeAnswerCountsAsAnsweredOnAForm()
        {
            var json = @"{""results"":[{""id"":""c1"",""type"":""course_reflection"",""section"":""cub"",""template"":""personal-development-course/cub"",""status"":""draft_review"",
                ""answers"":{""what_did_you_learn"":""Lots"",""date"":{""start"":""2026-01-02"",""end"":""2026-01-03""}}}]}";
            var templates = Templates();
            templates.Forms["personal-development-course/cub"] = ReadyParsers.ParseFormQuestions(@"{""document"":[
                {""id"":""date"",""type"":""date_range"",""label"":""Dates""},{""id"":""what_did_you_learn"",""type"":""text_input"",""label"":""Learn""}]}");
            var member = new ReadyMemberInput { MemberId = "a", MemberName = "Ann", Records = ReadyParsers.ParseAchievements(json) };

            var report = ReadyToSubmitBuilder.Build(new ReadyToSubmitInput { Section = "cub", Members = new List<ReadyMemberInput> { member }, Templates = templates });

            Assert.AreEqual("Personal Development Course", report.Items.Single().Badge);
        }

        [TestMethod]
        public void APlanIsJudgedByThePlanQuestionsAndAReviewByTheReviewQuestions()
        {
            var answered = new[] { ("project_name", "Compost"), ("Goal_vision", "Grow veg"), ("agree", "true") };
            var report = Build(
                Member("Ann", Record("special_interest_area", "draft_approval", "cub", "sia/cub/plan", answered)),
                Member("Ben", Record("special_interest_area", "draft_review", "cub", "sia/cub/plan", ("project_name", "Kayak"), ("highlight_SIA", "Finishing"))),
                Member("Cat", Record("special_interest_area", "draft_review", "cub", "sia/cub/plan", ("project_name", "Robot"))), // review not written
                Member("Dan", Record("special_interest_area", "in_progress", "cub", "sia/cub/plan", answered)), // plan already approved
                Member("Eve", Record("special_interest_area", "draft_approval", "cub", "sia/cub/plan", ("project_name", "Art"), ("Goal_vision", "Paint"), ("agree", "false")))); // box not ticked

            CollectionAssert.AreEqual(new[] { "Ann", "Ben" }, report.Items.Select(i => i.MemberName).ToList());
            Assert.AreEqual("SIA: Compost", report.Items[0].Badge);
            StringAssert.Contains(report.Items[0].Waiting, "approval");
            StringAssert.Contains(report.Items[1].Waiting, "review");
        }

        [TestMethod]
        public void PeakAwardPartsAreJudgedByTheirOwnTemplate()
        {
            var report = Build(
                Member("Ann", Record("course_reflection", "draft_review", "cub", "personal-development-course/cub", ("what_did_you_learn", "Lots"))),
                Member("Ben", Record("course_reflection", "draft_review", "cub", "personal-development-course/cub")),
                Member("Cat", Record("course_reflection", "draft_review", "cub", "personal-development-course/unknown", ("what_did_you_learn", "Lots")))); // template not read

            Assert.AreEqual("Ann", report.Items.Single().MemberName);
            Assert.AreEqual("Personal Development Course", report.Items.Single().Badge);
            Assert.AreEqual(ReadyCategory.PeakAward, report.Items.Single().Category);
        }

        [TestMethod]
        public void OnlyTheKindsOfBadgeAskedForAreListedAndTheOrderIsTheMembers()
        {
            var members = new[]
            {
                Member("Zed", Milestone("in_progress", "complete", 1)),
                Member("Amy", Milestone("in_progress", "complete", 2), Record("course_reflection", "draft_review", "cub", "personal-development-course/cub", ("what_did_you_learn", "x")))
            };

            var all = Build(members);
            CollectionAssert.AreEqual(new[] { "Zed", "Amy", "Amy" }, all.Items.Select(i => i.MemberName).ToList());

            var onlyMilestones = ReadyToSubmitBuilder.Build(new ReadyToSubmitInput
            {
                Section = "cub", Members = members.ToList(), Templates = Templates(), Categories = new List<ReadyCategory> { ReadyCategory.Milestone }
            });
            Assert.AreEqual(2, onlyMilestones.Items.Count);
            Assert.AreEqual(0, ReadyToSubmitBuilder.Build(new ReadyToSubmitInput { Section = "cub", Members = members.ToList(), Categories = new List<ReadyCategory>() }).Items.Count);
        }

        [TestMethod]
        public void OnlyWhatCouldBeWaitingNeedsATemplateRead()
        {
            Assert.IsTrue(ReadyToSubmitBuilder.NeedsFormTemplate(Record("special_interest_area", "draft_approval", "cub", "sia/cub/plan"), "cub"));
            Assert.IsFalse(ReadyToSubmitBuilder.NeedsFormTemplate(Record("special_interest_area", "awarded", "cub", "sia/cub/plan"), "cub"));
            Assert.IsFalse(ReadyToSubmitBuilder.NeedsFormTemplate(Record("special_interest_area", "draft_approval", "joey", "sia/joey/plan"), "cub"));
            Assert.AreEqual("sia/cub/review", ReadyToSubmitBuilder.FormTemplatePath(Record("special_interest_area", "draft_review", "cub", "sia/cub/plan")));
            Assert.AreEqual("sia/cub/plan", ReadyToSubmitBuilder.FormTemplatePath(Record("special_interest_area", "draft_approval", "cub", "sia/cub/plan")));
            Assert.IsFalse(ReadyToSubmitBuilder.NeedsOasTemplate(Oas("draft_review", "bushcraft", 3), "cub")); // nothing ticked off yet
            Assert.IsTrue(ReadyToSubmitBuilder.NeedsOasTemplate(Oas("draft_review", "bushcraft", 3, ("plan_one_verifiedDate", "2026-09-01")), "cub"));
        }

        [TestMethod]
        public void NobodyOrNothingWaitingStillBuilds()
        {
            Assert.AreEqual(0, Build().Items.Count);
            var report = Build(Member("Ann"));
            Assert.AreEqual(0, report.Items.Count);
            Assert.AreEqual(1, report.MembersChecked);
        }

        [TestMethod]
        public void TheUnitIsInPatrolOrder()
        {
            MemberListModel Scout(string first, string patrol, string duty = "") => new MemberListModel { id = first, first_name = first, last_name = "X", patrol_name = patrol, patrol_duty = duty };
            var ordered = ReadyToSubmitBuilder.InPatrolOrder(new[] { Scout("Zed", ""), Scout("Bea", "Eagles"), Scout("Dot", "Eagles", "PL"), Scout("Una", "Wolves", "UL"), Scout("Cal", "Eagles", "APL") });

            CollectionAssert.AreEqual(new[] { "Una", "Dot", "Cal", "Bea", "Zed" }, ordered.Select(m => m.first_name).ToList());
        }

        private static ReadyToSubmitReport Sample()
        {
            var report = Build(
                Member("Ann Smith", Milestone("in_progress", "complete", 1), Oas("draft_review", "bushcraft", 3, ("plan_one_verifiedDate", "2026-09-01"), ("do_one_verifiedDate", "2026-09-02"), ("logbook_up_to_date", "true"))),
                Member("Ben Jones", Record("special_interest_area", "draft_approval", "cub", "sia/cub/plan", ("project_name", "Compost"), ("Goal_vision", "Grow veg"), ("agree", "true"))),
                Member("Cat Nobody"));
            report.GeneratedAt = new DateTime(2026, 11, 7);
            return report;
        }

        private static ReadyToSubmitReport SampleWithEveryList()
        {
            var report = ReadyToSubmitBuilder.Build(new ReadyToSubmitInput
            {
                UnitName = "Cub Unit", Section = "cub", Templates = Templates(), LogbookTickOptional = true, MaxTicksShort = 2,
                Members = new List<ReadyMemberInput>
                {
                    Member("Ann Smith", Milestone("in_progress", "complete", 1), Camping("a", "b", "c")),
                    Member("Ben Jones", Camping("a", "b", "logbook_up_to_date")),
                    Member("Cat Brown", Camping("b", "c", "logbook_up_to_date")),
                    Member("Dan White", Camping("c", "logbook_up_to_date"))
                }
            });
            report.GeneratedAt = new DateTime(2026, 11, 7);
            return report;
        }

        [DataTestMethod]
        [DataRow(OutputType.PDF, "pdf")]
        [DataRow(OutputType.Excel, "xlsx")]
        public void TheReportRendersAsPdfAndExcelOnA4(OutputType outputType, string extension)
        {
            SetEnvironment.SetSyncfusionLicense();
            var request = new ReportGenerationRequest
            {
                ReportType = ReportType.ReadyToSubmit, GroupName = "Sample Group", Section = "cub", UnitName = "Cub Unit",
                OutputType = outputType, ReportData = JsonConvert.SerializeObject(Sample())
            };
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) };

            var result = new Function().FunctionHandler(proxyRequest, null);

            Assert.IsFalse(string.IsNullOrEmpty(result), $"{extension} came back empty");
            Directory.CreateDirectory("TestResults");
            File.WriteAllBytes($@"TestResults/ReadyToSubmit.{extension}", Convert.FromBase64String(result));
            if (outputType == OutputType.PDF)
            {
                using var document = new PdfLoadedDocument(Convert.FromBase64String(result));
                var text = document.Pages[0].ExtractText();
                StringAssert.Contains(text, "Ben Jones");
                StringAssert.Contains(text, "Bushcraft stage 3");
                Assert.AreEqual(595, Math.Min(document.Pages[0].Size.Width, document.Pages[0].Size.Height), 2);
            }
        }

        [TestMethod]
        public void TheLogbookAndNearlyThereListsEachGetAPageOfTheirOwn()
        {
            SetEnvironment.SetSyncfusionLicense();
            var request = new ReportGenerationRequest
            {
                ReportType = ReportType.ReadyToSubmit, GroupName = "Sample Group", Section = "cub", UnitName = "Cub Unit",
                OutputType = OutputType.PDF, ReportData = JsonConvert.SerializeObject(SampleWithEveryList())
            };
            var result = new Function().FunctionHandler(new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) }, null);

            Directory.CreateDirectory("TestResults");
            File.WriteAllBytes("TestResults/ReadyToSubmit_AllLists.pdf", Convert.FromBase64String(result));
            using var document = new PdfLoadedDocument(Convert.FromBase64String(result));
            Assert.AreEqual(3, document.Pages.Count);
            StringAssert.Contains(document.Pages[0].ExtractText(), "Milestone 1");
            StringAssert.Contains(document.Pages[1].ExtractText(), "Logbook");
            StringAssert.Contains(document.Pages[1].ExtractText(), "Ann Smith");
            StringAssert.Contains(document.Pages[2].ExtractText(), "Still to tick off");
            StringAssert.Contains(document.Pages[2].ExtractText(), "Item A");
            StringAssert.Contains(document.Pages[2].ExtractText(), "Camping stage 2");
            StringAssert.Contains(document.Pages[2].ExtractText(), "Ben Jones");
        }

        [TestMethod]
        public void AnEmptyReportRendersToo()
        {
            SetEnvironment.SetSyncfusionLicense();
            var empty = Build(Member("Cat Nobody"));
            var request = new ReportGenerationRequest
            {
                ReportType = ReportType.ReadyToSubmit, GroupName = "Sample Group", Section = "cub", UnitName = "Cub Unit",
                OutputType = OutputType.PDF, ReportData = JsonConvert.SerializeObject(empty)
            };
            var result = new Function().FunctionHandler(new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) }, null);

            Assert.IsFalse(string.IsNullOrEmpty(result));
            using var document = new PdfLoadedDocument(Convert.FromBase64String(result));
            StringAssert.Contains(document.Pages[0].ExtractText(), "Nothing is ready");
        }
    }
}
