using Topo.Model.Milestone;

namespace Topo.Model.Wallchart
{
    // Years of service is tracked per section (from the first event the member was credited with in the section, not from
    // their intro to the section or their overall Scouting membership), and only counts whole years actually completed -
    // a member gets the "2 years" badge once they pass their second anniversary in the section, not on their first day of it.
    public static class YearsOfServiceCalculator
    {
        // The date of the earliest event that earned the member a credit (as a participant or a leader) in the section, from the
        // event logs on their milestone records. The records Terrain returns include the member's earlier sections, so the
        // caller passes only the section's own. Nothing when no event is logged, for example a member whose progress was
        // all imported, or who has no milestone record yet.
        public static DateTime? FirstEvent(IEnumerable<MilestoneResult> sectionMilestones)
        {
            DateTime? first = null;
            foreach (var record in sectionMilestones)
            {
                foreach (var logged in record.event_log)
                {
                    if (logged.event_start_datetime == default)
                        continue;
                    // Terrain's times are UTC, so an early morning event belongs to the day before there; the date that counts is the local one
                    var start = logged.event_start_datetime;
                    var date = (start.Kind == DateTimeKind.Utc ? start.ToLocalTime() : start).Date;
                    if (first == null || date < first)
                        first = date;
                }
            }
            return first;
        }

        // Whole years since the member's first event in the section. A member with no logged event falls back to their intro
        // to the section, so a long-standing member whose events Terrain did not log is not shown with no service at all.
        public static int? CalculateCompletedYears(DateTime? firstEvent, DateTime? introToSection, DateTime asOf) =>
            CalculateCompletedYears(firstEvent ?? introToSection, asOf);

        public static int? CalculateCompletedYears(DateTime? startOfService, DateTime asOf)
        {
            if (!startOfService.HasValue)
                return null;

            var years = asOf.Year - startOfService.Value.Year;
            if (startOfService.Value.AddYears(years) > asOf)
                years--;

            return years < 0 ? 0 : years;
        }
    }
}
