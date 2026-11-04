namespace Topo.Model.Wallchart
{
    public class WallchartItemModel
    {
        public string MemberName { get; set; } = string.Empty;
        public string MemberId { get; set; } = string.Empty;
        public string MemberPatrol { get; set; } = string.Empty;
        public DateTime? IntroToScouting { get; set; }
        public DateTime? IntroToSection { get; set; }
        public int Milestone1Community { get; set; }
        public int Milestone1Outdoors { get; set; }
        public int Milestone1Creative { get; set; }
        public int Milestone1PersonalGrowth { get; set; }
        public int Milestone1Assist { get; set; }
        public int Milestone1Lead { get; set; }
        public DateTime? Milestone1Awarded { get; set; }
        public DateTime? Milestone1Presented { get; set; }
        // True when Milestone 1 is marked "not required" in Terrain - the counts above are 0 in
        // that case too, so this is how ComputePeakAwardPercent's Milestones bucket tells "skipped"
        // apart from "not started yet".
        public bool Milestone1Skipped { get; set; }
        public int Milestone2Community { get; set; }
        public int Milestone2Outdoors { get; set; }
        public int Milestone2Creative { get; set; }
        public int Milestone2PersonalGrowth { get; set; }
        public int Milestone2Assist { get; set; }
        public int Milestone2Lead { get; set; }
        public DateTime? Milestone2Awarded { get; set; }
        public DateTime? Milestone2Presented { get; set; }
        public bool Milestone2Skipped { get; set; }
        public int Milestone3Community { get; set; }
        public int Milestone3Outdoors { get; set; }
        public int Milestone3Creative { get; set; }
        public int Milestone3PersonalGrowth { get; set; }
        public int Milestone3Assist { get; set; }
        public int Milestone3Lead { get; set; }
        public DateTime? Milestone3Awarded { get; set; }
        public DateTime? Milestone3Presented { get; set; }
        public bool Milestone3Skipped { get; set; }
        public int OASBushcraftStage { get; set; }
        public int OASBushwalkingStage { get; set; }
        public int OASCampingStage { get; set; }
        public int OASAlpineStage { get; set; }
        public int OASCyclingStage { get; set; }
        public int OASVerticalStage { get; set; }
        public int OASAquaticsStage { get; set; }
        public int OASBoatingStage { get; set; }
        public int OASPaddlingStage { get; set; }
        public int NightsCamped { get; set; }
        public int NightsCampedInSection { get; set; }
        public int KMsHiked { get; set; }
        public int OASStageProgressions { get; set; }
        public int SIAAdventureSport { get; set; }
        public int SIAArtsLiterature { get; set; }
        public int SIAEnvironment { get; set; }
        public int SIAStemInnovation { get; set; }
        public int SIAGrowthDevelopment { get; set; }
        public int SIACreatingABetterWorld { get; set; }
        public DateTime? LeadershipCourse { get; set; }
        public DateTime? AdventurousJourney { get; set; }
        public DateTime? PersonalReflection { get; set; }
        // Fraction (0.0-1.0) complete towards this member's section's own top award - see
        // WallchartService.ComputePeakAwardPercent for how it's derived. Only meaningful until
        // the award itself is granted, at which point PeakAwardAwarded takes over instead.
        public double PeakAward { get; set; }
        public DateTime? PeakAwardAwarded { get; set; }

    }
}
