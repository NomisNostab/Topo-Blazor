using Topo.Model.Milestone;

namespace Topo.Model.Progress
{
    // Terrain can hold several records for the same achievement (an Introduction to Section entered twice, a milestone
    // that exists as both in progress and awarded, the same Outdoor Adventure Skill stage awarded twice). These pick the
    // record to use, and count each stage once, in a way that doesn't depend on the order Terrain returns them in.
    public static class AchievementSelection
    {
        // The awarded Introduction to Scouting / Section entered most recently. For the section one, only that section's.
        public static Result? LatestAwardedIntro(IEnumerable<Result> results, string? section = null) =>
            results
                .Where(r => r.status == "awarded" && (section == null || SameSection(r.section, section)))
                .OrderByDescending(r => r.status_updated)
                .FirstOrDefault();

        // The awarded Peak Award record (or one of its parts) entered most recently in the section
        public static PeakAwardResult? LatestAwardedPeak(IEnumerable<PeakAwardResult> results, string section) =>
            results
                .Where(r => r.status == "awarded" && SameSection(r.section, section))
                .OrderByDescending(r => r.status_updated)
                .FirstOrDefault();

        // One milestone record per stage and section, in stage order. A member can earn Milestone 1 in each section, so
        // different sections' milestones are never treated as duplicates of each other. An awarded record beats one
        // that is not, and among records of the same kind the most recently entered wins.
        public static List<MilestoneResult> OnePerMilestoneStage(IEnumerable<MilestoneResult> results) =>
            results
                .GroupBy(r => (stage: r.achievement_meta.stage, section: (r.section ?? "").Trim().ToLowerInvariant()))
                .OrderBy(g => g.Key.stage)
                .ThenBy(g => g.Key.section)
                .Select(g => g
                    .OrderByDescending(r => r.status == "awarded")
                    .ThenByDescending(r => r.status_updated)
                    .First())
                .ToList();

        // Outdoor Adventure Skill progressions made in a section. Each stream and stage counts once however many
        // records Terrain holds for it, and a stage of 7 or more counts as two progressions.
        public static int OasProgressions(IEnumerable<OAS.Result> results, string section, bool includeApproved)
        {
            var stages = results
                .Where(r => (r.status == "awarded" || (includeApproved && r.status == "approved")) && SameSection(r.section, section))
                .Select(r => (r.achievement_meta.stream, r.achievement_meta.stage))
                .Distinct()
                .ToList();
            return stages.Count + stages.Count(s => s.stage >= 7);
        }

        private static bool SameSection(string? recordSection, string section) =>
            string.Equals(recordSection, section, StringComparison.OrdinalIgnoreCase);
    }
}
