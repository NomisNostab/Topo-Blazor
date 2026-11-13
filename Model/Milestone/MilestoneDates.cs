namespace Topo.Model.Milestone
{
    public static class MilestoneDates
    {
        // When a milestone was awarded: Terrain's own date, else the imported one, else when the status last changed.
        // The rule is shared by every kind of award, see AwardDate.
        public static DateTime AwardedDate(MilestoneResult result) =>
            AwardDate.Resolve(result.date_awarded, result.imported?.date_awarded, result.status_updated);
    }
}
