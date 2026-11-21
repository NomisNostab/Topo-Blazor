using Topo.Model.Approvals;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class SubmissionDetailsTests
    {
        private const string Template = """
        { "template": "sia/scout/plan", "version": 2, "document": [
          { "title": "Project Details", "input_groups": [
            { "title": "Pick one", "inputs": [
              { "id": "area", "type": "single_tab_selector", "label": "Which area?", "selections": [ { "id": "sia_a", "title": "Area A" }, { "id": "sia_b", "title": "Area B" } ] } ] },
            { "inputs": [ { "id": "name", "type": "text_input", "label": "Name your project" } ] },
            { "inputs": [ { "id": "empty", "type": "text_input", "label": "Not answered" } ] }
          ] },
          { "title": "Goals", "input_groups": [
            { "inputs": [
              { "id": "sdg", "type": "radio", "label": "Connected?", "selections": [ { "id": "yes_id", "title": "Yes" }, { "id": "no_id", "title": "No" } ] },
              { "id": "which", "type": "multi_tab_selector", "label": "Which?", "tabs": [
                  { "title": "People", "selections": [ { "id": "g1", "title": "1. No Poverty" } ] },
                  { "title": "Planet", "selections": [ { "id": "g2", "title": "2. Climate" } ] } ] },
              { "id": "note", "type": "markdown_text", "label": "Remember to plan" },
              { "id": "files", "type": "file_uploader", "label": "Upload Files" } ] } ] }
        ] }
        """;

        private const string Achievement = """
        { "template": "sia/scout/plan", "version": 2, "answers": {
            "area": "sia_b", "name": "Write a book", "sdg": "yes_id", "which": [ "g1", "g2" ],
            "files": [ "id-a", "id-b" ] },
          "uploads": [ { "id": "id-a", "filename": "a.pdf", "uploaded_on": "2026-11-20T01:00:00+00:00" }, { "id": "id-b", "filename": "b.png" }, { "id": "id-other", "filename": "other.txt" } ] }
        """;

        private static List<DetailField> Fields(SubmissionDetails details) =>
            details.Sections.SelectMany(s => s.Groups).SelectMany(g => g.Fields).ToList();

        [TestMethod]
        public void SectionsGroupsAndQuestionsFollowTheTemplate()
        {
            var details = SubmissionDetails.Build(Template, Achievement)!;
            CollectionAssert.AreEqual(new[] { "Project Details", "Goals" }, details.Sections.Select(s => s.Title).ToArray());
            Assert.AreEqual("Pick one", details.Sections[0].Groups[0].Title);
            CollectionAssert.AreEqual(new[] { "Which area?", "Name your project", "Not answered", "Connected?", "Which?", "Remember to plan", "Upload Files" },
                Fields(details).Select(f => f.Label).ToArray());
        }

        [TestMethod]
        public void TextAnswersAreShownAndOptionsAreShownByTheirTitleNotTheirId()
        {
            var fields = Fields(SubmissionDetails.Build(Template, Achievement)!);
            Assert.AreEqual("Area B", fields.Single(f => f.Label == "Which area?").Value);
            Assert.AreEqual("Write a book", fields.Single(f => f.Label == "Name your project").Value);
            Assert.AreEqual("Yes", fields.Single(f => f.Label == "Connected?").Value);
            Assert.AreEqual("1. No Poverty, 2. Climate", fields.Single(f => f.Label == "Which?").Value);
        }

        [TestMethod]
        public void AQuestionWithNoAnswerIsStillListedWithNothingInIt()
        {
            var field = Fields(SubmissionDetails.Build(Template, Achievement)!).Single(f => f.Label == "Not answered");
            Assert.AreEqual("", field.Value);
        }

        [TestMethod]
        public void NotesFromTheTemplateAreKeptAsNotes()
        {
            var field = Fields(SubmissionDetails.Build(Template, Achievement)!).Single(f => f.Label == "Remember to plan");
            Assert.AreEqual(DetailFieldKind.Note, field.Kind);
        }

        [TestMethod]
        public void AnUploadQuestionShowsTheFilesItsAnswerNamesWithTheirDates()
        {
            var files = Fields(SubmissionDetails.Build(Template, Achievement)!).Single(f => f.Kind == DetailFieldKind.Files).Files;
            CollectionAssert.AreEqual(new[] { "a.pdf", "b.png" }, files.Select(f => f.FileName).ToArray());
            Assert.IsNotNull(files[0].UploadedOn);
            Assert.IsNull(files[1].UploadedOn);
        }

        [TestMethod]
        public void AnUploadQuestionWithNoAnswerShowsNoFiles()
        {
            var achievement = """{ "answers": { "name": "x" }, "uploads": [ { "id": "id-a", "filename": "a.pdf" } ] }""";
            Assert.AreEqual(0, Fields(SubmissionDetails.Build(Template, achievement)!).Single(f => f.Kind == DetailFieldKind.Files).Files.Count);
        }

        private const string OasTemplate = """
        { "document": [ { "title": "Stage 5 - Camping", "input_groups": [ { "title": "Plan>", "inputs": [
            { "id": "req_one", "type": "verifiable_checkbox", "label": "I can plan." },
            { "id": "req_two", "type": "verifiable_checkbox", "label": "I can cook." },
            { "id": "req_three", "type": "verifiable_checkbox", "label": "I can pack." },
            { "id": "logbook", "type": "checkbox", "label": "My logbook is up to date." } ] } ] } ] }
        """;

        [TestMethod]
        public void TickedRequirementsShowWhoVerifiedThemAndWhen()
        {
            var achievement = """
            { "answers": {
                "req_one": true, "req_one_verifierName": "Bart Williams, Susana Prat", "req_one_verifierContact": "N/A", "req_one_verifiedDate": "08/10/2024",
                "req_two": true, "req_two_verifierName": "Scout Leaders", "req_two_verifierContact": "", "req_two_verifiedDate": "18/11/2026",
                "req_three": false, "logbook": true } }
            """;
            var fields = Fields(SubmissionDetails.Build(OasTemplate, achievement)!);

            Assert.IsTrue(fields[0].Ticked);
            Assert.AreEqual("Verified by Bart Williams, Susana Prat (N/A). Date 08/10/2024", fields[0].Verification);
            Assert.AreEqual("Verified by Scout Leaders. Date 18/11/2026", fields[1].Verification);
            Assert.IsFalse(fields[2].Ticked);
            Assert.AreEqual("", fields[2].Verification);
            Assert.IsTrue(fields[3].Ticked);
            Assert.AreEqual("", fields[3].Verification);
        }

        [TestMethod]
        public void ARecordWithNoAnswersShowsTheFormUnticked()
        {
            var fields = Fields(SubmissionDetails.Build(OasTemplate, """{ "template": "x" }""")!);
            Assert.AreEqual(4, fields.Count);
            Assert.IsTrue(fields.All(f => !f.Ticked));
        }

        [TestMethod]
        public void UnreadableTemplatesOrRecordsGiveNothingToShow()
        {
            Assert.IsNull(SubmissionDetails.Build(null, Achievement));
            Assert.IsNull(SubmissionDetails.Build(Template, ""));
            Assert.IsNull(SubmissionDetails.Build("not json", Achievement));
            Assert.IsNull(SubmissionDetails.Build("""{ "template": "x" }""", Achievement));
            Assert.IsNull(SubmissionDetails.Build(Template, "[]"));
        }
    }
}
