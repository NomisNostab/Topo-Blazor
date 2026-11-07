namespace Topo.Model.SiaProgress
{
    // Which projects a report covers
    public enum SiaScope
    {
        OneProject,
        OnePerson,
        WholeUnit
    }

    // One question of a project's plan or review, from Terrain's template, with the wording shown to the Scout
    public class SiaQuestion
    {
        public string Id { get; set; } = string.Empty;
        public string Label { get; set; } = string.Empty;
    }

    public class SiaTemplate
    {
        public List<SiaQuestion> Questions { get; set; } = new List<SiaQuestion>();
        // What a radio or tab option stands for, by the id Terrain stores in the answer
        public Dictionary<string, string> Options { get; set; } = new Dictionary<string, string>();
    }

    // The two parts of a project, each with its own questions
    public class SiaTemplates
    {
        public SiaTemplate Plan { get; set; } = new SiaTemplate();
        public SiaTemplate Review { get; set; } = new SiaTemplate();
    }

    // A question as Terrain words it, with what the Scout wrote
    public class SiaAnswer
    {
        public string Question { get; set; } = string.Empty;
        public string Answer { get; set; } = string.Empty;
    }

    // A Scout's Special Interest Area project, to talk through with them on one page
    public class SiaProject
    {
        public string MemberId { get; set; } = string.Empty;
        public string MemberName { get; set; } = string.Empty;
        public string ProjectName { get; set; } = string.Empty;
        public string Area { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public bool IsComplete { get; set; }
        public DateTime StatusUpdated { get; set; }
        public DateTime? DateAwarded { get; set; }
        public List<SiaAnswer> Plan { get; set; } = new List<SiaAnswer>();
        public List<SiaAnswer> Review { get; set; } = new List<SiaAnswer>();
        public List<string> Files { get; set; } = new List<string>();
    }

    // Where one Scout is up to against what the section's top award asks for
    public class SiaSummaryRow
    {
        public string MemberName { get; set; } = string.Empty;
        public int Completed { get; set; }
        public int InProgress { get; set; }
        public int AreasCompleted { get; set; }
    }

    public class SiaProgressReport
    {
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public bool IncludeSummary { get; set; }
        // Each Scout's pages (and the summary) take whole sheets of paper, with a blank page where one is needed
        public bool DoubleSided { get; set; }
        // What the section's top award asks for; 0 when the section isn't known
        public string AwardName { get; set; } = string.Empty;
        public int SiaRequired { get; set; }
        public int AreasRequired { get; set; }
        public List<SiaSummaryRow> Summary { get; set; } = new List<SiaSummaryRow>();
        public List<SiaProject> Projects { get; set; } = new List<SiaProject>();

        // The summary's wording, so the page and the exported report say the same
        public string CompletedText(SiaSummaryRow row) => SiaRequired > 0 ? $"{row.Completed}/{SiaRequired}" : row.Completed.ToString();

        public string AreasText(SiaSummaryRow row) => AreasRequired > 0 ? $"{row.AreasCompleted}/{AreasRequired}" : row.AreasCompleted.ToString();

        public string SummaryNote() => SiaRequired > 0
            ? $"The {AwardName} asks for {SiaRequired} Special Interest Area projects over at least {AreasRequired} areas. Completed projects are awarded ones; in progress is every other project started in this section."
            : "Completed projects are awarded ones; in progress is every other project started in this section.";
    }
}
