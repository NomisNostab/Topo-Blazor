namespace Topo.Model.Milestone
{
    public class MilestoneSummaryListModel
    {
        public string memberName { get; set; } = string.Empty;
        public int currentLevel { get; set; }
        public int percentComplete { get; set; }
        // True once Milestone 3 itself is awarded/approved - there's no "Milestone 4" to move on to,
        // so currentLevel stays at 3 either way. The report uses this to split "still working on
        // Milestone 3" from "has completed Milestone 3" into separate sections.
        public bool milestone3Awarded { get; set; }
        public int milestone1ParticipateCommunity { get; set; }
        public int milestone1ParticipateOutdoors { get; set; }
        public int milestone1ParticipateCreative { get; set; }
        public int milestone1ParticipatePersonalGrowth { get; set; }
        public int milestone1Assist { get; set; }
        public string milestone1AssistArea { get; set; } = string.Empty;
        public int milestone1Lead { get; set; }
        public string milestone1LeadArea { get; set; } = string.Empty; 
        public int milestone2ParticipateCommunity { get; set; }
        public int milestone2ParticipateOutdoors { get; set; }
        public int milestone2ParticipateCreative { get; set; }
        public int milestone2ParticipatePersonalGrowth { get; set; }
        public int milestone2Assist { get; set; }
        public string milestone2AssistArea { get; set; } = string.Empty;
        public int milestone2Lead { get; set; }
        public string milestone2LeadArea { get; set; } = string.Empty;
        public int milestone3ParticipateCommunity { get; set; }
        public int milestone3ParticipateOutdoors { get; set; }
        public int milestone3ParticipateCreative { get; set; }
        public int milestone3ParticipatePersonalGrowth { get; set; }
        public int milestone3Assist { get; set; }
        public string milestone3AssistArea { get; set; } = string.Empty;
        public int milestone3Lead { get; set; }
        public string milestone3LeadArea { get; set; } = string.Empty;

    }
}
