using System.Globalization;

namespace Topo.Model
{
    // Terrain can hold an award's date in three places, and which one is filled in depends on how the award was recorded:
    //   date_awarded            at the top of the record, for awards recorded in Terrain
    //   imported.date_awarded   for awards imported from before Terrain
    //   status_updated          always there, but it is when the record last changed, not necessarily when it was awarded
    // So the date to show is the first of those that exists, in that order.
    public static class AwardDate
    {
        // The date Terrain recorded for the award, or null when it only has status_updated.
        // Terrain writes these dates as yyyy-mm-dd; a missing or odd value counts as "not there".
        public static DateTime? Recorded(string? dateAwarded, string? importedDateAwarded) =>
            Parse(dateAwarded) ?? Parse(importedDateAwarded);

        public static DateTime Resolve(string? dateAwarded, string? importedDateAwarded, DateTime statusUpdated) =>
            Recorded(dateAwarded, importedDateAwarded) ?? statusUpdated;

        private static DateTime? Parse(string? text) =>
            DateTime.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date) ? date : null;
    }
}
