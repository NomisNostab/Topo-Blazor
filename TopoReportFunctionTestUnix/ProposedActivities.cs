using Topo.Model.Program;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class ProposedActivitiesTests
    {
        private const string One = @"{
            ""id"": ""p-1"", ""title"": ""Night Hike"", ""description"": ""Walk with torches"", ""justification"": ""We have not done one"",
            ""leader"": ""Me and the leaders"", ""challenge_area"": ""personal_growth"", ""created_at"": ""2026-08-26T01:02:03.000Z"",
            ""proposing_member"": { ""id"": ""m-1"", ""first_name"": ""Sam"", ""last_name"": ""Example"", ""patrol_name"": ""Lyrebird"" }, ""uploads"": [] }";

        private static string Item(string title, string area = "outdoors", string created = "2026-01-01T00:00:00.000Z", string by = "A B") =>
            $@"{{ ""id"": ""{title}"", ""title"": ""{title}"", ""description"": """", ""justification"": """", ""leader"": """", ""challenge_area"": ""{area}"", ""created_at"": ""{created}"",
                ""proposing_member"": {{ ""first_name"": ""{by.Split(' ')[0]}"", ""last_name"": ""{by.Split(' ')[1]}"", ""patrol_name"": """" }}, ""uploads"": [] }}";

        [TestMethod]
        public void AProposalIsReadWithEverythingTheYouthWrote()
        {
            var list = ProposedActivities.Parse("[" + One + "]")!;
            var p = list.Single();
            Assert.AreEqual("p-1", p.Id);
            Assert.AreEqual("Night Hike", p.Title);
            Assert.AreEqual("Walk with torches", p.Description);
            Assert.AreEqual("We have not done one", p.Justification);
            Assert.AreEqual("Me and the leaders", p.SuggestedLead);
            Assert.AreEqual("Personal Growth", p.ChallengeArea);
            Assert.AreEqual("Sam Example", p.ProposedBy);
            Assert.AreEqual("Lyrebird", p.Patrol);
            // the time Terrain gives is UTC; it is shown as a local date, so allow for the zone this runs in
            Assert.AreEqual(new DateTimeOffset(2026, 8, 26, 1, 2, 3, TimeSpan.Zero).LocalDateTime, p.Submitted);
        }

        [TestMethod]
        public void TheListCanComeAloneOrInsideAnObject()
        {
            Assert.AreEqual(1, ProposedActivities.Parse("[" + One + "]")!.Count);
            Assert.AreEqual(1, ProposedActivities.Parse(@"{ ""results"": [" + One + "] }")!.Count);
            Assert.AreEqual(1, ProposedActivities.Parse(@"{ ""proposals"": [" + One + "] }")!.Count);
            Assert.AreEqual(1, ProposedActivities.Parse(@"{ ""something"": 1, ""other"": [" + One + "] }")!.Count);
        }

        [TestMethod]
        public void NothingReadableIsNullAndNoProposalsIsAnEmptyList()
        {
            Assert.IsNull(ProposedActivities.Parse(""));
            Assert.IsNull(ProposedActivities.Parse("   "));
            Assert.IsNull(ProposedActivities.Parse("not json"));
            Assert.IsNull(ProposedActivities.Parse(@"{ ""message"": ""Unauthorized"" }"));
            Assert.IsNull(ProposedActivities.Parse("42"));
            Assert.AreEqual(0, ProposedActivities.Parse("[]")!.Count);
            Assert.AreEqual(0, ProposedActivities.Parse(@"{ ""results"": [] }")!.Count);
        }

        [TestMethod]
        public void MissingPiecesDoNotStopAProposalShowing()
        {
            var list = ProposedActivities.Parse(@"[ { ""title"": ""Bare"" }, { ""title"": ""  "" }, { ""id"": ""x"" }, ""text"", { ""title"": ""No member"", ""challenge_area"": ""creative"", ""created_at"": ""soon"" } ]")!;
            CollectionAssert.AreEqual(new[] { "Bare", "No member" }, list.Select(p => p.Title).ToList());
            Assert.AreEqual("", list[0].ProposedBy);
            Assert.AreEqual("Not applicable", list[0].ChallengeArea);
            Assert.IsNull(list[0].Submitted);
            Assert.IsNull(list[1].Submitted);
            Assert.AreEqual("Creative", list[1].ChallengeArea);
        }

        [TestMethod]
        public void ChallengeAreasAreNamedLikeTerrainShowsThem()
        {
            Assert.AreEqual("Outdoors", ProposedActivities.AreaLabel("outdoors"));
            Assert.AreEqual("Creative", ProposedActivities.AreaLabel("creative"));
            Assert.AreEqual("Community", ProposedActivities.AreaLabel("community"));
            Assert.AreEqual("Personal Growth", ProposedActivities.AreaLabel("personal_growth"));
            Assert.AreEqual("Not applicable", ProposedActivities.AreaLabel("not_applicable"));
            Assert.AreEqual("Not applicable", ProposedActivities.AreaLabel(""));
            Assert.AreEqual("new_area", ProposedActivities.AreaLabel("new_area"));
        }

        [TestMethod]
        public void ProposalsSortByTheChosenColumnWithTheTitleBreakingTies()
        {
            var list = ProposedActivities.Parse("[" + string.Join(",", new[]
            {
                Item("Zip Line", "outdoors", "2026-03-01T00:00:00.000Z", "Cal Dee"),
                Item("Art Night", "creative", "2026-01-01T00:00:00.000Z", "Bo Eve"),
                Item("Bake Sale", "community", "2026-02-01T00:00:00.000Z", "Al Fox"),
                Item("Abseil", "outdoors", "2026-02-01T00:00:00.000Z", "Al Fox")
            }) + "]")!;

            CollectionAssert.AreEqual(new[] { "Art Night", "Abseil", "Bake Sale", "Zip Line" }, Titles(list, ProposedActivitySortColumn.Submitted, false));
            CollectionAssert.AreEqual(new[] { "Zip Line", "Bake Sale", "Abseil", "Art Night" }, Titles(list, ProposedActivitySortColumn.Submitted, true));
            CollectionAssert.AreEqual(new[] { "Abseil", "Art Night", "Bake Sale", "Zip Line" }, Titles(list, ProposedActivitySortColumn.Activity, false));
            CollectionAssert.AreEqual(new[] { "Bake Sale", "Art Night", "Abseil", "Zip Line" }, Titles(list, ProposedActivitySortColumn.ChallengeArea, false));
            CollectionAssert.AreEqual(new[] { "Abseil", "Bake Sale", "Art Night", "Zip Line" }, Titles(list, ProposedActivitySortColumn.ProposedBy, false));
        }

        [TestMethod]
        public void AProposalWithoutAnIdIsStillToldApart()
        {
            var withId = new ProposedActivityModel { Id = "p-1", Title = "Origami" };
            Assert.AreEqual("p-1", withId.OpenKey);

            var first = new ProposedActivityModel { Title = "Origami", Submitted = new DateTime(2026, 1, 1) };
            var second = new ProposedActivityModel { Title = "Origami", Submitted = new DateTime(2026, 2, 1) };
            var other = new ProposedActivityModel { Title = "Egg Blowing", Submitted = new DateTime(2026, 1, 1) };
            Assert.AreNotEqual(first.OpenKey, second.OpenKey);
            Assert.AreNotEqual(first.OpenKey, other.OpenKey);
            // and with no date either, it still has a key
            Assert.AreEqual("Origami|", new ProposedActivityModel { Title = "Origami" }.OpenKey);
        }

        [TestMethod]
        public void SortingNothingGivesNothing()
        {
            Assert.AreEqual(0, ProposedActivities.Sort(new List<ProposedActivityModel>(), ProposedActivitySortColumn.Activity, true).Count);
        }

        private static ProposedActivityModel Idea(string title) => new ProposedActivityModel { Title = title };

        private static bool Same(string first, string second) => ProposedActivities.SameActivity(ProposedActivities.NameWords(first), ProposedActivities.NameWords(second));

        [TestMethod]
        public void NamesAreReducedToTheirMainWords()
        {
            CollectionAssert.AreEquivalent(new[] { "capture", "flag" }, ProposedActivities.NameWords("  Capture   The Flag ").ToList());
            CollectionAssert.AreEquivalent(new[] { "paper", "plane" }, ProposedActivities.NameWords("Paper Planes Night").ToList());
            CollectionAssert.AreEquivalent(new[] { "bushwalk" }, ProposedActivities.NameWords("Bush-walk!".Replace("-", "")).ToList());
            CollectionAssert.AreEquivalent(new[] { "egg", "blowing" }, ProposedActivities.NameWords("egg, blowing").ToList());
            // a name that is all filler is still compared by what it has
            CollectionAssert.AreEquivalent(new[] { "night" }, ProposedActivities.NameWords("Night").ToList());
            Assert.AreEqual(0, ProposedActivities.NameWords("").Count);
            Assert.AreEqual(0, ProposedActivities.NameWords(null!).Count);
            // words that only end in s are not made shorter when that would leave too little, or when it is part of the word
            CollectionAssert.AreEquivalent(new[] { "bus", "glass" }, ProposedActivities.NameWords("Bus Glass").ToList());
        }

        [TestMethod]
        public void SimilarNamesAreTheSameActivity()
        {
            Assert.IsTrue(Same("Capture The Flag", "capture the flag "));
            Assert.IsTrue(Same("Origami", "Origami Night"));
            Assert.IsTrue(Same("Origami Night", "Origami"));
            Assert.IsTrue(Same("Paper Plane Night", "Paper Planes"));
            Assert.IsTrue(Same("Bushwalk", "Fun Overnight Bushwalk"));
            Assert.IsTrue(Same("Rock Climbing", "Rock climbing at Mt Tinbeerwah"));
        }

        [TestMethod]
        public void DifferentActivitiesAreNotTheSame()
        {
            Assert.IsFalse(Same("Hall Games", "Water Games"));
            Assert.IsFalse(Same("Origami", "Paper Plane Night"));
            Assert.IsFalse(Same("Capture The Flag", "Silent Ball"));
            // a name with nothing to compare never matches, not even another with nothing
            Assert.IsFalse(Same("", "Origami"));
            Assert.IsFalse(Same("", ""));
        }

        [TestMethod]
        public void SimilarlyNamedEventsAreCountedAsRunsAndListed()
        {
            var now = new DateTime(2026, 11, 1);
            var proposals = new List<ProposedActivityModel> { Idea("Origami"), Idea("Hall Games") };
            var events = new[]
            {
                ("Origami Night", new DateTime(2025, 3, 5)),
                ("Origami Night", new DateTime(2025, 6, 5)),
                ("Origami", new DateTime(2026, 8, 1)),
                ("Water Games", new DateTime(2026, 1, 1)),
                ("Hall games", new DateTime(2024, 1, 1))
            };

            ProposedActivities.AddHistory(proposals, events, now);

            Assert.AreEqual(3, proposals[0].TimesRun);
            Assert.AreEqual(new DateTime(2026, 8, 1), proposals[0].LastRun);
            CollectionAssert.AreEqual(new[] { "Origami Night (2)", "Origami" }, proposals[0].RunTitles);
            Assert.AreEqual(1, proposals[1].TimesRun);
            CollectionAssert.AreEqual(new[] { "Hall games" }, proposals[1].RunTitles);
        }

        [TestMethod]
        public void SimilarProposalsAreCountedTogether()
        {
            var proposals = new List<ProposedActivityModel> { Idea("Origami"), Idea("Origami Night"), Idea("Capture The Flag") };
            ProposedActivities.AddHistory(proposals, null, new DateTime(2026, 11, 1));
            CollectionAssert.AreEqual(new[] { 2, 2, 1 }, proposals.Select(p => p.TimesProposed).ToList());
        }

        [TestMethod]
        public void HowOftenAnActivityWasProposedAndRunIsCounted()
        {
            var now = new DateTime(2026, 11, 1);
            var proposals = new List<ProposedActivityModel> { Idea("Capture The Flag"), Idea("capture the flag "), Idea("Origami"), Idea("Rock Climbing") };
            var events = new[]
            {
                ("Capture the Flag", new DateTime(2025, 3, 5)),
                ("Capture the flag", new DateTime(2026, 5, 20)),
                ("Origami", new DateTime(2024, 9, 1)),
                ("Hall Games", new DateTime(2026, 1, 1)),
                // not run yet, so not run before
                ("Rock Climbing", new DateTime(2026, 12, 1))
            };

            ProposedActivities.AddHistory(proposals, events, now);

            CollectionAssert.AreEqual(new[] { 2, 2, 1, 1 }, proposals.Select(p => p.TimesProposed).ToList());
            CollectionAssert.AreEqual(new int?[] { 2, 2, 1, 0 }, proposals.Select(p => p.TimesRun).ToList());
            Assert.AreEqual(new DateTime(2026, 5, 20), proposals[0].LastRun);
            Assert.AreEqual(new DateTime(2024, 9, 1), proposals[2].LastRun);
            Assert.IsNull(proposals[3].LastRun);
        }

        [TestMethod]
        public void WithoutEventsTheTimesRunAreUnknownNotNone()
        {
            var proposals = new List<ProposedActivityModel> { Idea("Origami"), Idea("Origami") };
            ProposedActivities.AddHistory(proposals, null, new DateTime(2026, 11, 1));
            Assert.IsTrue(proposals.All(p => p.TimesRun == null && p.LastRun == null));
            Assert.IsTrue(proposals.All(p => p.TimesProposed == 2));

            // no proposals and no events are fine too
            ProposedActivities.AddHistory(new List<ProposedActivityModel>(), new (string, DateTime)[0], new DateTime(2026, 11, 1));
            var none = new List<ProposedActivityModel> { Idea("Origami") };
            ProposedActivities.AddHistory(none, new (string, DateTime)[0], new DateTime(2026, 11, 1));
            Assert.AreEqual(0, none[0].TimesRun);
        }

        [TestMethod]
        public void ProposalsSortByHowOftenProposedAndRun()
        {
            var a = new ProposedActivityModel { Title = "A", TimesProposed = 3, TimesRun = null };
            var b = new ProposedActivityModel { Title = "B", TimesProposed = 1, TimesRun = 4 };
            var c = new ProposedActivityModel { Title = "C", TimesProposed = 2, TimesRun = 0 };
            var list = new List<ProposedActivityModel> { a, b, c };
            CollectionAssert.AreEqual(new[] { "B", "C", "A" }, Titles(list, ProposedActivitySortColumn.TimesProposed, false));
            CollectionAssert.AreEqual(new[] { "A", "C", "B" }, Titles(list, ProposedActivitySortColumn.TimesRun, false));
        }

        private static List<string> Titles(List<ProposedActivityModel> list, ProposedActivitySortColumn column, bool descending) =>
            ProposedActivities.Sort(list, column, descending).Select(p => p.Title).ToList();
    }
}
