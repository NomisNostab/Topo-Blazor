using Newtonsoft.Json;
using Syncfusion.Pdf.Parsing;
using Topo.Model.Members;
using Topo.Model.ReportGeneration;
using Topo.Model.SIA;
using Topo.Model.SiaProgress;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class SiaProgress
    {
        private const string PlanTemplate = @"{""template"":""sia/scout/plan"",""version"":2,""document"":[
            {""id"":""special_interest_area_selection"",""type"":""single_tab_selector"",""label"":""Identify the Special Interest Area*"",
              ""selections"":[{""id"":""sia_environment"",""title"":""Environment""},{""id"":""sia_stem_innovation"",""title"":""STEM & Innovation""}]},
            {""id"":""project_name"",""type"":""text_input"",""label"":""Name your project*""},
            {""id"":""Goal_vision"",""type"":""text_input"",""label"":""My goal is...""},
            {""id"":""who_is_involved"",""type"":""radio"",""label"":""Who is involved in your project?"",
              ""selections"":[{""id"":""who_involved_other"",""title"":""Other""},{""id"":""who_involved_project_individual"",""title"":""Project Patrol""}]},
            {""id"":""which_goals"",""type"":""multi_tab_selector"",""label"":""If so, which one(s)?"",
              ""tabs"":[{""id"":""goals_sustainable_zero_hunger"",""title"":""2. Zero Hunger""},{""id"":""goals_sustainable_climate_action"",""title"":""13. Climate Action""}]},
            {""id"":""attach_documentation"",""type"":""file_uploader"",""label"":""Do you have any files to attach here?""}]}";

        private const string ReviewTemplate = @"{""document"":[{""id"":""highlight_SIA"",""type"":""text_input"",""label"":""What was the highlight of this Special Interest Area?""},
            {""id"":""continue_extend_SIA"",""type"":""text_input"",""label"":""How could you continue or extend this Special Interest Area?""}]}";

        private static SiaTemplates Templates() => new SiaTemplates
        {
            Plan = SiaTemplateParser.Parse(PlanTemplate),
            Review = SiaTemplateParser.Parse(ReviewTemplate)
        };

        private static Result Project(string id, string name, string area, string status, string section = "scout", int month = 3) => new Result
        {
            id = id, section = section, status = status, status_updated = new DateTime(2026, month, 10),
            answers = new Answers { project_name = name, special_interest_area_selection = area }
        };

        private static SiaMemberInput Member(string name, params Result[] projects) =>
            new SiaMemberInput { MemberId = name, MemberName = name, Results = new GetSIAResultsModel { results = projects } };

        private static SiaProgressInput Input() => new SiaProgressInput
        {
            UnitName = "Scout Unit", Section = "scout", Templates = Templates(),
            AwardName = "Australian Scout Award", SiaRequired = 6, AreasRequired = 3,
            Members = new List<SiaMemberInput>
            {
                Member("Sam Sample",
                    Project("p1", "Billy cart", "sia_adventure_sport", "awarded"),
                    Project("p2", "Compost", "sia_environment", "awarded", month: 4),
                    Project("p3", "Robot", "sia_stem_innovation", "in_progress", month: 5),
                    Project("p4", "Cub project", "sia_environment", "awarded", section: "cub")),
                Member("Alex Another", Project("p5", "Kayak", "sia_adventure_sport", "pending_review")),
                Member("Pat Nothing")
            }
        };

        [TestMethod]
        public void TheTemplateGivesTheQuestionsInOrderAndTheWordingOfOptions()
        {
            var template = SiaTemplateParser.Parse(PlanTemplate);

            CollectionAssert.AreEqual(
                new[] { "special_interest_area_selection", "project_name", "Goal_vision", "who_is_involved", "which_goals" },
                template.Questions.Select(q => q.Id).ToList()); // not the file uploader
            Assert.AreEqual("Name your project", template.Questions[1].Label); // without the required marker
            Assert.AreEqual("Project Patrol", template.Options["who_involved_project_individual"]);
            Assert.AreEqual("STEM & Innovation", template.Options["sia_stem_innovation"]);
            Assert.AreEqual(0, SiaTemplateParser.Parse("not json").Questions.Count);
            Assert.AreEqual(0, SiaTemplateParser.Parse("{}").Questions.Count);
        }

        [TestMethod]
        public void AllTheProjectsOfTheUnitInThisSectionOnly()
        {
            var report = SiaProgressBuilder.Build(Input());

            CollectionAssert.AreEqual(new[] { "Billy cart", "Compost", "Robot", "Kayak" }, report.Projects.Select(p => p.ProjectName).ToList());
            Assert.AreEqual("STEM & Innovation", report.Projects.Single(p => p.ProjectName == "Robot").Area);
        }

        [TestMethod]
        public void CompletedProjectsCanBeLeftOutButNotWhenAskedForOneProject()
        {
            var input = Input();
            input.ExcludeCompleted = true;
            CollectionAssert.AreEqual(new[] { "Robot", "Kayak" }, SiaProgressBuilder.Build(input).Projects.Select(p => p.ProjectName).ToList());

            input.Members = input.Members.Take(1).ToList();
            input.ProjectId = "p1"; // an awarded one, asked for by name
            Assert.AreEqual("Billy cart", SiaProgressBuilder.Build(input).Projects.Single().ProjectName);
        }

        [TestMethod]
        public void TheSummaryCountsCompletedInProgressAndAreasAgainstTheAward()
        {
            var input = Input();
            input.IncludeSummary = true;
            input.ExcludeCompleted = true; // the summary is about everything, whatever the pages show

            var report = SiaProgressBuilder.Build(input);

            Assert.AreEqual(6, report.SiaRequired);
            Assert.AreEqual(3, report.AreasRequired);
            var sam = report.Summary.Single(s => s.MemberName == "Sam Sample");
            Assert.AreEqual(2, sam.Completed); // the cub one is another section's
            Assert.AreEqual(1, sam.InProgress);
            Assert.AreEqual(2, sam.AreasCompleted);
            var alex = report.Summary.Single(s => s.MemberName == "Alex Another");
            Assert.AreEqual(0, alex.Completed);
            Assert.AreEqual(1, alex.InProgress);
            var pat = report.Summary.Single(s => s.MemberName == "Pat Nothing");
            Assert.AreEqual(0, pat.Completed + pat.InProgress + pat.AreasCompleted);
            CollectionAssert.AreEqual(new[] { "Sam Sample", "Alex Another", "Pat Nothing" }, report.Summary.Select(s => s.MemberName).ToList()); // as given
        }

        [TestMethod]
        public void NoSummaryUnlessAskedFor()
        {
            Assert.AreEqual(0, SiaProgressBuilder.Build(Input()).Summary.Count);
        }

        [TestMethod]
        public void AnswersAreShownWithTheQuestionAndTheWordingOfWhatWasPicked()
        {
            var input = Input();
            var project = input.Members[0].Results.results.Single(r => r.id == "p1");
            project.answers.Goal_vision = "Race it";
            project.answers.who_is_involved = "who_involved_project_individual";
            project.answers.which_goals = new[] { "goals_sustainable_zero_hunger", "goals_sustainable_climate_action" };
            project.answers.highlight_SIA = "Finishing";
            project.uploads = new[] { new Upload { filename = "cart.jpg" } };
            project.imported = new Imported { date_awarded = "2025-11-02" };

            var built = SiaProgressBuilder.Build(input).Projects.Single(p => p.ProjectName == "Billy cart");

            CollectionAssert.AreEqual(new[] { "My goal is...", "Who is involved in your project?", "If so, which one(s)?" }, built.Plan.Select(a => a.Question).ToList());
            Assert.AreEqual("Project Patrol", built.Plan[1].Answer);
            Assert.AreEqual("2. Zero Hunger, 13. Climate Action", built.Plan[2].Answer);
            Assert.AreEqual("Finishing", built.Review[0].Answer);
            Assert.AreEqual("", built.Review[1].Answer);
            CollectionAssert.AreEqual(new[] { "cart.jpg" }, built.Files);
            Assert.AreEqual(new DateTime(2025, 11, 2), built.DateAwarded); // the imported date, not when it was touched
            Assert.AreEqual("Awarded", built.Status);
            Assert.IsTrue(built.IsComplete);
        }

        [TestMethod]
        public void QuestionsThatOnlySometimesApplyAreLeftOutWhenBlank()
        {
            var built = SiaProgressBuilder.Build(Input()).Projects.Single(p => p.ProjectName == "Kayak");

            Assert.IsFalse(built.Plan.Any(a => a.Question == "If so, which one(s)?"));
            Assert.IsTrue(built.Plan.Any(a => a.Question == "My goal is..." && a.Answer == ""));
            Assert.AreEqual("Pending review", built.Status);
            Assert.IsNull(built.DateAwarded);
        }

        [TestMethod]
        public void AQuestionTheReportCantReadAnAnswerForIsLeftOut()
        {
            var input = Input();
            input.Templates.Plan.Questions.Add(new SiaQuestion { Id = "a_new_question", Label = "Something Terrain added" });

            var built = SiaProgressBuilder.Build(input).Projects.First();

            Assert.IsFalse(built.Plan.Any(a => a.Question == "Something Terrain added"));
        }

        [TestMethod]
        public void AnUnansweredQuestionIsLeftBlankToWriteIn()
        {
            var input = new SiaProgressInput { UnitName = "Scout Unit", Section = "scout", Templates = Templates() };
            input.Members.Add(Member("Alex Another", Project("p5", "Kayak", "sia_adventure_sport", "pending_review")));
            var base64 = Render(SiaProgressBuilder.Build(input), OutputType.PDF, "SiaProgress_Blank.pdf");
            using var document = new PdfLoadedDocument(Convert.FromBase64String(base64));
            var text = document.Pages[0].ExtractText();

            StringAssert.Contains(text, "My goal is..."); // the question is there
            Assert.IsFalse(text.Contains("Not answered"), "nothing is written where there is no answer");
        }

        [TestMethod]
        public void NobodyOrNoProjectsStillBuilds()
        {
            var empty = new SiaProgressInput { Section = "scout", IncludeSummary = true };
            var report = SiaProgressBuilder.Build(empty);
            Assert.AreEqual(0, report.Projects.Count + report.Summary.Count);

            empty.Section = "";
            empty.Members.Add(Member("Pat Nothing"));
            Assert.AreEqual(1, SiaProgressBuilder.Build(empty).Summary.Count);
        }

        [TestMethod]
        public void TheUnitIsInPatrolOrderLikeTheSignInSheet()
        {
            MemberListModel Scout(string first, string patrol, string duty = "") =>
                new MemberListModel { id = first, first_name = first, last_name = "X", patrol_name = patrol, patrol_duty = duty };
            var members = new[]
            {
                Scout("Zed", ""), Scout("Bea", "Eagles"), Scout("Cal", "Eagles", "APL"), Scout("Dot", "Eagles", "PL"),
                Scout("Eve", "Wolves", "PL"), Scout("Una", "Wolves", "UL"), Scout("Abe", "Wolves")
            };

            var ordered = SiaProgressBuilder.InPatrolOrder(members).Select(m => m.first_name).ToList();

            // the unit leader, then Eagles (PL, APL, the rest), Wolves, then anyone without a patrol
            CollectionAssert.AreEqual(new[] { "Una", "Dot", "Cal", "Bea", "Eve", "Abe", "Zed" }, ordered);
        }

        private static SiaProgressInput ThreeScouts(bool summary, bool doubleSided)
        {
            var input = new SiaProgressInput { UnitName = "Scout Unit", Section = "scout", Templates = Templates(), IncludeSummary = summary, DoubleSided = doubleSided };
            input.Members.Add(Member("Scout A", Project("a1", "A one", "sia_environment", "awarded")));
            input.Members.Add(Member("Scout B", Project("b1", "B one", "sia_environment", "awarded"), Project("b2", "B two", "sia_stem_innovation", "in_progress")));
            input.Members.Add(Member("Scout C", Project("c1", "C one", "sia_environment", "awarded"), Project("c2", "C two", "sia_stem_innovation", "in_progress"), Project("c3", "C three", "sia_adventure_sport", "pending_review")));
            return input;
        }

        private static int Pages(SiaProgressInput input)
        {
            var base64 = Render(SiaProgressBuilder.Build(input), OutputType.PDF, "SiaProgress_Pages.pdf");
            using var document = new PdfLoadedDocument(Convert.FromBase64String(base64));
            return document.Pages.Count;
        }

        [TestMethod]
        public void DoubleSidedGivesEachScoutWholeSheetsOfPaper()
        {
            // a page for each project: 1 + 2 + 3, and the summary
            Assert.AreEqual(6, Pages(ThreeScouts(summary: false, doubleSided: false)));
            Assert.AreEqual(7, Pages(ThreeScouts(summary: true, doubleSided: false)));
            // A's one page gets a blank back, B's two are a whole sheet, C's three start on a new sheet (the last needs no padding)
            Assert.AreEqual(7, Pages(ThreeScouts(summary: false, doubleSided: true)));
            // the summary has a blank back too
            Assert.AreEqual(9, Pages(ThreeScouts(summary: true, doubleSided: true)));
        }

        private static SiaProgressInput WithALongProject(bool doubleSided)
        {
            var input = new SiaProgressInput { UnitName = "Scout Unit", Section = "scout", Templates = Templates(), DoubleSided = doubleSided };
            var wordy = string.Join(" ", Enumerable.Repeat("This is a long answer that goes on and on about the project.", 15));
            var long1 = Project("l1", "A very big project", "sia_environment", "awarded");
            long1.answers.Goal_vision = wordy;
            long1.answers.highlight_SIA = wordy;
            long1.answers.continue_extend_SIA = wordy;
            input.Members.Add(Member("Wordy Scout", long1));
            input.Members.Add(Member("Short Scout", Project("s1", "Small", "sia_environment", "awarded")));
            input.Members.Add(Member("Another Scout", Project("s2", "Another", "sia_stem_innovation", "in_progress")));
            return input;
        }

        [TestMethod]
        public void AProjectTooLongForAPageRunsOntoASecondAndTheSecondNamesWhoseItIs()
        {
            var base64 = Render(SiaProgressBuilder.Build(WithALongProject(doubleSided: false)), OutputType.PDF, "SiaProgress_Long.pdf");
            using var document = new PdfLoadedDocument(Convert.FromBase64String(base64));

            Assert.AreEqual(4, document.Pages.Count); // 2 + 1 + 1
            StringAssert.Contains(document.Pages[1].ExtractText(), "Wordy Scout");
        }

        [TestMethod]
        public void DoubleSidedPaddingCountsTheExtraPage()
        {
            // 2 pages (a whole sheet, no blank), then 1, then a blank before the next Scout, then 1
            Assert.AreEqual(5, Pages(WithALongProject(doubleSided: true)));
        }

        private static string Render(SiaProgressReport report, OutputType outputType, string name)
        {
            SetEnvironment.SetSyncfusionLicense();
            var request = new ReportGenerationRequest
            {
                ReportType = ReportType.SiaProgress, GroupName = "Sample Group", Section = "scout", UnitName = "Scout Unit",
                OutputType = outputType, ReportData = JsonConvert.SerializeObject(report)
            };
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest { Body = JsonConvert.SerializeObject(request) };

            var result = new Function().FunctionHandler(proxyRequest, null);

            Assert.IsFalse(string.IsNullOrEmpty(result), $"{name} came back empty");
            Directory.CreateDirectory("TestResults");
            File.WriteAllBytes($@"TestResults/{name}", Convert.FromBase64String(result));
            return result;
        }

        private static SiaProgressInput Rich()
        {
            var input = Input();
            var project = input.Members[0].Results.results.Single(r => r.id == "p1");
            project.answers.Goal_vision = "To build a cart that is safe, steers well and finishes the course.";
            project.answers.who_is_involved = "who_involved_other";
            project.answers.highlight_SIA = "Crossing the finish line on race day.";
            project.uploads = new[] { new Upload { filename = "billy-cart-design.jpg" } };
            input.IncludeSummary = true;
            return input;
        }

        [DataTestMethod]
        [DataRow(OutputType.PDF, "pdf")]
        [DataRow(OutputType.Excel, "xlsx")]
        public void ReportsRenderWithASummaryAndOnePagePerProject(OutputType outputType, string extension)
        {
            Render(SiaProgressBuilder.Build(Rich()), outputType, $"SiaProgress.{extension}");
        }

        [TestMethod]
        public void RendersWhenThereAreNoProjectsOnlyASummaryOrNothing()
        {
            var input = Rich();
            input.Members.ForEach(m => m.Results = new GetSIAResultsModel());
            Render(SiaProgressBuilder.Build(input), OutputType.PDF, "SiaProgress_SummaryOnly.pdf");

            input.IncludeSummary = false;
            Render(SiaProgressBuilder.Build(input), OutputType.PDF, "SiaProgress_Nothing.pdf");
            Render(SiaProgressBuilder.Build(input), OutputType.Excel, "SiaProgress_Nothing.xlsx");
        }

        [TestMethod]
        public void ALargeUnitStaysWithinWhatTheReportFunctionCanReturn()
        {
            var input = Rich();
            for (var i = 0; i < 12; i++)
                input.Members.Add(Member($"Scout Number{i}", Project($"a{i}", $"Project {i} with a fairly long name", "sia_environment", "awarded"),
                    Project($"b{i}", $"Second {i}", "sia_stem_innovation", "in_progress"), Project($"c{i}", $"Third {i}", "sia_adventure_sport", "pending_review")));

            var report = SiaProgressBuilder.Build(input);

            // Lambda answers up to 6 MB (a little under that once encoded)
            foreach (var (outputType, name) in new[] { (OutputType.PDF, "SiaProgress_Large.pdf"), (OutputType.Excel, "SiaProgress_Large.xlsx") })
            {
                var result = Render(report, outputType, name);
                Assert.IsTrue(result.Length < 6 * 1024 * 1024 * 0.9, $"the {name} response was {result.Length / 1024} KB");
            }
        }
    }
}
