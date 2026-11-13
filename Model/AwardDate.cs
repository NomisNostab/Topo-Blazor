using System.Globalization;

namespace Topo.Model
{
    // Terrain can hold an award's date in three places, and which one is filled in depends on how the award was recorded:
    //   date_awarded            at the top of the record, for awards recorded in Terrain
    //   imported.date_awarded   for awards imported from before Terrain
    //   status_updated          always there, but it is when the record last changed, not necessarily when it was awarded
    // So the date to show is the first of those that exists, in that order.
    //
    // One wrinkle: Terrain works out date_awarded in UTC, so an award made early in the morning in Australia is dated the
    // day before the user made it (until 8am in Western Australia, and until between 9:30am and 11am in the other states
    // and territories, depending on daylight saving). status_updated is a moment in time that the app reads in the timezone of the user's own device,
    // so wherever they are in Australia, when date_awarded is the UTC day of status_updated the day to show is the
    // user's own day for that moment. When it is a different day (the award was dated on purpose, or the record was
    // touched later) Terrain's date is used as it is.
    //
    // The result is a day as the user sees it, with no timezone attached (DateTimeKind.Unspecified), like the dates Terrain
    // writes. The reports are made by a function that runs in UTC, and a date that still carried the user's offset would
    // be moved back a day when it got there.
    public static class AwardDate
    {
        public static DateTime Resolve(string? dateAwarded, string? importedDateAwarded, DateTime statusUpdated)
        {
            // Terrain writes these dates as yyyy-mm-dd; a missing or odd value counts as "not there"
            var terrainsDate = Parse(dateAwarded);
            if (terrainsDate == null)
                return Parse(importedDateAwarded) ?? DateTime.SpecifyKind(statusUpdated, DateTimeKind.Unspecified);

            return terrainsDate.Value == statusUpdated.ToUniversalTime().Date
                ? DateTime.SpecifyKind(statusUpdated.Date, DateTimeKind.Unspecified)
                : terrainsDate.Value;
        }

        private static DateTime? Parse(string? text) =>
            DateTime.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date) ? date : null;
    }
}
