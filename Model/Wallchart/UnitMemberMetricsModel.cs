namespace Topo.Model.Wallchart
{
    // One member's row from Terrain's unit metrics, the call behind the table on Terrain's Group Life page.
    // Only the part the Wallchart reads is listed; Terrain sends a lot more per member.
    public class UnitMemberMetricsModel
    {
        public string member_id { get; set; } = string.Empty;
        public UnitMemberAchievementMetricsModel achievement_metrics { get; set; } = new UnitMemberAchievementMetricsModel();
    }

    public class UnitMemberAchievementMetricsModel
    {
        // Terrain's own "Peak Award Progress", 0-100. Can carry a fraction (38.5); Terrain shows it as a whole percent.
        public double peak_award_progress { get; set; }
    }
}
