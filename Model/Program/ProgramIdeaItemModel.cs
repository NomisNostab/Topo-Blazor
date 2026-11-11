namespace Topo.Model.Program
{
    public class ProgramIdeaItemModel
    {
        public string ActivityName { get; set; } = string.Empty;
        public string ChallengeArea { get; set; } = string.Empty;
        public DateTime LastRunDate { get; set; }
        public int TimesRun { get; set; }
        public bool Highlight { get; set; }
    }
}
