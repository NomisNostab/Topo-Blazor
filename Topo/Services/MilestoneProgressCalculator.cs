namespace Topo.Services
{
    // The per-milestone percent-complete formula, shared so the dedicated Milestone report
    // (MilestoneService) and the Wallchart's Peak Award Milestones bucket (WallchartService)
    // can never show a different percentage for the same underlying progress.
    public static class MilestoneProgressCalculator
    {
        public static int CalculatePercentComplete(int level, int participantCommunity, int participantOutdoors, int participantCreative, int participantPersonalGrowth, int leaderTotal)
        {
            // Milestone 3 caps each participate category at 4, and leadership at 6, regardless of
            // how many were actually logged - matches the original per-milestone report formula.
            int participantTotal = level == 3
                ? Math.Min(4, participantCommunity) + Math.Min(4, participantOutdoors) + Math.Min(4, participantCreative) + Math.Min(4, participantPersonalGrowth)
                : participantCommunity + participantOutdoors + participantCreative + participantPersonalGrowth;
            var total = level == 3 ? participantTotal + Math.Min(6, leaderTotal) : participantTotal + leaderTotal;

            var target = level switch
            {
                1 => 26,
                2 => 24,
                3 => 20,
                _ => 1
            };

            return total * 100 / target;
        }
    }
}
