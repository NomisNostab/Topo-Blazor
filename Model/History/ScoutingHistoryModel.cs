namespace Topo.Model.History
{
    // One member's whole scouting history, split by section, with totals across all sections.
    public class ScoutingHistoryModel
    {
        public string MemberName { get; set; } = string.Empty;
        public DateTime? ScoutingStart { get; set; }
        public string TimeInScouting { get; set; } = string.Empty;
        public List<SectionHistory> Sections { get; set; } = new List<SectionHistory>();
        public HistoryTotals Totals { get; set; } = new HistoryTotals();
    }

    public class SectionHistory
    {
        // joey, cub, scout, venturer, rover
        public string SectionKey { get; set; } = string.Empty;
        // Joeys, Cubs, Scouts...
        public string SectionName { get; set; } = string.Empty;
        // The unit(s) the member's submissions were made under, e.g. "Cub Unit"
        public string Units { get; set; } = string.Empty;
        public DateTime? From { get; set; }
        public DateTime? To { get; set; }
        public string TimeInSection { get; set; } = string.Empty;

        public List<MilestoneEntry> Milestones { get; set; } = new List<MilestoneEntry>();
        public string FinalMilestone { get; set; } = string.Empty;
        public List<OasStreamEntry> Oas { get; set; } = new List<OasStreamEntry>();
        public int OasProgressions { get; set; }
        public List<SiaEntry> Sia { get; set; } = new List<SiaEntry>();
        public int SiaInProgress { get; set; }
        public List<AwardEntry> AdditionalAwards { get; set; } = new List<AwardEntry>();

        public int Nights { get; set; }
        public int Camps { get; set; }
        public int Hikes { get; set; }
        public double HikeKm { get; set; }

        // Roles on logbook activities
        public int Led { get; set; }
        public int Assisted { get; set; }
        public int Participated { get; set; }
        public string CurrentRole { get; set; } = string.Empty;
    }

    public class MilestoneEntry
    {
        public int Stage { get; set; }
        public string Status { get; set; } = string.Empty;
        public DateTime? Date { get; set; }
    }

    public class OasStreamEntry
    {
        public string Stream { get; set; } = string.Empty;
        // The highest stage reached in this stream by the end of the section
        public int FinalStage { get; set; }
        // Stages 7 to 9 count as two progressions each
        public int Progressions { get; set; }
        // The stages achieved in this section
        public List<OasStageEntry> Stages { get; set; } = new List<OasStageEntry>();
    }

    public class OasStageEntry
    {
        public int Stage { get; set; }
    }

    public class SiaEntry
    {
        public string Area { get; set; } = string.Empty;
        public string Project { get; set; } = string.Empty;
        public DateTime? Date { get; set; }
    }

    public class AwardEntry
    {
        public string Name { get; set; } = string.Empty;
        public DateTime? Date { get; set; }
    }

    public class HistoryTotals
    {
        public int Milestones { get; set; }
        public int OasProgressions { get; set; }
        public int Sia { get; set; }
        public int AdditionalAwards { get; set; }
        public int Nights { get; set; }
        public int Camps { get; set; }
        public int Hikes { get; set; }
        public double HikeKm { get; set; }
        public int Led { get; set; }
        public int Assisted { get; set; }
        public int Participated { get; set; }
    }

    // How a member's roles on logged activities read. Assists are no longer a role in Terrain, so "assisted" only
    // appears for a member who still has some from before.
    public static class ScoutingHistoryRoles
    {
        public static string Text(int led, int assisted, int participated) =>
            assisted > 0
                ? $"Led {led}, assisted {assisted}, participated in {participated} logged activities"
                : $"Led {led}, participated in {participated} logged activities";
    }
}
