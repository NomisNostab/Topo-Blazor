using System.Globalization;

namespace Topo.Model.Milestone
{
    public static class MilestoneDates
    {
        // When a milestone was awarded. One brought across from before Terrain carries its real date in imported.date_awarded
        // (status_updated is then only the day it was imported); otherwise it is when the status last changed.
        public static DateTime AwardedDate(MilestoneResult result)
        {
            return DateTime.TryParseExact(result.imported?.date_awarded, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var imported)
                ? imported
                : result.status_updated;
        }
    }
}
