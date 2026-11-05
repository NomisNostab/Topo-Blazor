using System.Globalization;
using Topo.Model.Logbook;
using Topo.Model.Progress;
using AdditionalAwardResults = Topo.Model.AdditionalAwards.GetUnitAchievementsResultModel;
using MilestoneResults = Topo.Model.Milestone.GetMilestoneResultModel;
using OasResults = Topo.Model.OAS.GetUnitAchievementsResultsModel;
using SiaResults = Topo.Model.SIA.GetSIAResultsModel;

namespace Topo.Model.History
{
    // The unit a record was submitted under is only on the all-achievements list (latest_submission.unit_id)
    public class GetMemberAchievementsResultModel
    {
        public MemberAchievement[] results { get; set; } = new MemberAchievement[0];
    }

    public class MemberAchievement
    {
        public string id { get; set; } = string.Empty;
        public string section { get; set; } = string.Empty;
        public LatestSubmission? latest_submission { get; set; }
    }

    // A unit as listed for a group (members/groups/{id}/units)
    public class GroupUnit
    {
        public string id { get; set; } = string.Empty;
        public string name { get; set; } = string.Empty;
        public string section { get; set; } = string.Empty;
        public string group_id { get; set; } = string.Empty;
    }

    public class LatestSubmission
    {
        public string unit_id { get; set; } = string.Empty;
    }

    public class ScoutingHistoryInput
    {
        public string MemberName { get; set; } = string.Empty;
        public DateTime AsOf { get; set; } = DateTime.Today;
        public MilestoneResults Milestones { get; set; } = new MilestoneResults();
        public OasResults Oas { get; set; } = new OasResults();
        public SiaResults Sia { get; set; } = new SiaResults();
        public AdditionalAwardResults AdditionalAwards { get; set; } = new AdditionalAwardResults();
        public GetIntroductionResultsModel IntroToScouting { get; set; } = new GetIntroductionResultsModel();
        public GetIntroductionResultsModel IntroToSection { get; set; } = new GetIntroductionResultsModel();
        public GetMemberAchievementsResultModel AllAchievements { get; set; } = new GetMemberAchievementsResultModel();
        public List<GetMemberLogbookDetailResultModel> Logbook { get; set; } = new List<GetMemberLogbookDetailResultModel>();
        // unit id -> name, for the units in the member's group that we can look up
        public Dictionary<string, string> UnitNames { get; set; } = new Dictionary<string, string>();
        // additional award id -> title
        public Dictionary<string, string> AwardNames { get; set; } = new Dictionary<string, string>();
        // the member's current patrol/unit role, if known
        public string CurrentRole { get; set; } = string.Empty;
        public string CurrentSection { get; set; } = string.Empty;
    }

    public static class ScoutingHistoryBuilder
    {
        private static readonly string[] SectionOrder = { "joey", "cub", "scout", "venturer", "rover" };

        public static ScoutingHistoryModel Build(ScoutingHistoryInput input)
        {
            var model = new ScoutingHistoryModel { MemberName = input.MemberName };

            var sectionKeys = input.Milestones.results.Select(r => r.section)
                .Concat(input.Oas.results.Select(r => r.section))
                .Concat(input.Sia.results.Select(r => r.section))
                .Concat(input.AdditionalAwards.results.Select(r => r.section))
                .Concat(input.IntroToSection.results.Select(r => r.section))
                .Concat(input.AllAchievements.results.Select(r => r.section))
                .Select(Key)
                .Where(k => k != "")
                .Distinct()
                .OrderBy(SectionRank)
                .ThenBy(k => k)
                .ToList();

            // When each section started: the Introduction to Section, else the earliest dated record there
            var starts = new Dictionary<string, DateTime?>();
            foreach (var key in sectionKeys)
            {
                var intro = input.IntroToSection.results
                    .Where(r => Key(r.section) == key && r.status == "awarded")
                    .Select(r => (DateTime?)AwardedDate(r.imported.date_awarded, r.status_updated))
                    .OrderBy(d => d).FirstOrDefault();
                starts[key] = intro ?? EarliestRecordDate(input, key);
            }

            // Section i runs until the next section starts; the last runs to now
            var sections = new List<SectionHistory>();
            for (var i = 0; i < sectionKeys.Count; i++)
            {
                var key = sectionKeys[i];
                var next = sectionKeys.Skip(i + 1).Select(k => starts[k]).FirstOrDefault(d => d != null);
                var section = new SectionHistory
                {
                    SectionKey = key,
                    SectionName = SectionName(key),
                    From = starts[key],
                    To = starts[key] == null ? null : next ?? input.AsOf,
                    Units = UnitLabel(input, key)
                };
                section.TimeInSection = Duration(section.From, section.To);
                FillBadges(input, section, sectionKeys.Take(i + 1).ToList());
                sections.Add(section);
            }

            AssignLogbook(input, sections);

            if (!string.IsNullOrEmpty(input.CurrentSection))
            {
                var current = sections.FirstOrDefault(s => s.SectionKey == Key(input.CurrentSection));
                if (current != null)
                    current.CurrentRole = input.CurrentRole;
            }

            model.Sections = sections;
            model.ScoutingStart = input.IntroToScouting.results.Where(r => r.status == "awarded")
                .Select(r => (DateTime?)AwardedDate(r.imported.date_awarded, r.status_updated))
                .OrderBy(d => d).FirstOrDefault()
                ?? sections.Select(s => s.From).Where(d => d != null).OrderBy(d => d).FirstOrDefault();
            model.TimeInScouting = Duration(model.ScoutingStart, input.AsOf);
            model.Totals = new HistoryTotals
            {
                Milestones = sections.Sum(s => s.Milestones.Count(m => m.Status == "Awarded")),
                OasProgressions = sections.Sum(s => s.OasProgressions),
                Sia = sections.Sum(s => s.Sia.Count),
                AdditionalAwards = sections.Sum(s => s.AdditionalAwards.Count),
                Nights = sections.Sum(s => s.Nights),
                Camps = sections.Sum(s => s.Camps),
                Hikes = sections.Sum(s => s.Hikes),
                HikeKm = Math.Round(sections.Sum(s => s.HikeKm), 1),
                Led = sections.Sum(s => s.Led),
                Assisted = sections.Sum(s => s.Assisted),
                Participated = sections.Sum(s => s.Participated)
            };
            return model;
        }

        private static void FillBadges(ScoutingHistoryInput input, SectionHistory section, List<string> sectionsUpToHere)
        {
            var key = section.SectionKey;

            section.Milestones = input.Milestones.results
                .Where(r => Key(r.section) == key)
                .OrderBy(r => r.achievement_meta.stage)
                .Select(r => new MilestoneEntry
                {
                    Stage = r.achievement_meta.stage,
                    Status = StatusText(r.status),
                    Date = r.status == "awarded" ? AwardedDate(r.imported.date_awarded, r.status_updated) : null
                })
                .ToList();
            var finalMilestone = section.Milestones.Where(m => m.Status == "Awarded").Select(m => m.Stage).DefaultIfEmpty(0).Max();
            section.FinalMilestone = finalMilestone == 0 ? "None awarded" : $"Milestone {finalMilestone}";

            // OAS: per stream, the stages reached in this section; the final level counts the earlier sections too
            bool Achieved(string status) => status == "awarded" || status == "approved";
            var achievedOas = input.Oas.results.Where(r => Achieved(r.status)).ToList();
            foreach (var stream in achievedOas.Where(r => Key(r.section) == key).Select(r => r.achievement_meta.stream).Distinct().OrderBy(s => s))
            {
                var inSection = achievedOas.Where(r => Key(r.section) == key && r.achievement_meta.stream == stream)
                    .OrderBy(r => r.achievement_meta.stage).ToList();
                section.Oas.Add(new OasStreamEntry
                {
                    Stream = Title(stream),
                    FinalStage = achievedOas
                        .Where(r => r.achievement_meta.stream == stream && sectionsUpToHere.Contains(Key(r.section)))
                        .Max(r => r.achievement_meta.stage),
                    // stages 7, 8 and 9 count as two progressions each
                    Progressions = inSection.Sum(r => r.achievement_meta.stage >= 7 ? 2 : 1),
                    Stages = inSection.Select(r => new OasStageEntry { Stage = r.achievement_meta.stage }).ToList()
                });
            }
            section.OasProgressions = section.Oas.Sum(o => o.Progressions);

            var sia = input.Sia.results.Where(r => Key(r.section) == key).ToList();
            section.Sia = sia.Where(r => r.status == "awarded")
                .OrderBy(r => AwardedDate(r.imported.date_awarded, r.status_updated))
                .Select(r => new SiaEntry
                {
                    Area = Title((r.answers.special_interest_area_selection ?? "").Replace("sia_", "")),
                    Project = r.answers.project_name ?? "",
                    Date = AwardedDate(r.imported.date_awarded, r.status_updated)
                }).ToList();
            section.SiaInProgress = sia.Count(r => r.status != "awarded");

            section.AdditionalAwards = input.AdditionalAwards.results
                .Where(r => Key(r.section) == key && r.status == "awarded")
                .OrderBy(r => AwardedDate(r.imported.date_awarded, r.status_updated))
                .Select(r => new AwardEntry
                {
                    Name = input.AwardNames.TryGetValue(r.achievement_meta.additional_award_id, out var name)
                        ? name
                        : Title(r.achievement_meta.additional_award_id),
                    Date = AwardedDate(r.imported.date_awarded, r.status_updated)
                }).ToList();
        }

        // Each logbook entry belongs to the section the member was in on that date
        private static void AssignLogbook(ScoutingHistoryInput input, List<SectionHistory> sections)
        {
            var dated = sections.Where(s => s.From != null).OrderBy(s => s.From).ToList();
            if (dated.Count == 0)
                dated = sections.Take(1).ToList();
            if (dated.Count == 0)
                return;

            foreach (var entry in input.Logbook)
            {
                if (!DateTime.TryParseExact(entry.start_date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start))
                    continue;
                DateTime.TryParseExact(entry.end_date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end);

                var section = dated.LastOrDefault(s => s.From <= start) ?? dated.First();
                var categories = entry.categories ?? new string[0];
                if (categories.Contains("camping"))
                {
                    section.Camps++;
                    section.Nights += end > start ? (int)(end - start).TotalDays : 0;
                }
                if (categories.Contains("walking_hike"))
                {
                    section.Hikes++;
                    section.HikeKm += entry.distance_walkabout / 1000.0;
                }

                var role = (entry.details?.your_role ?? "").Trim().ToLowerInvariant();
                if (role.StartsWith("lead"))
                    section.Led++;
                else if (role.Contains("assist"))
                    section.Assisted++;
                else
                    section.Participated++;
            }

            foreach (var section in sections)
                section.HikeKm = Math.Round(section.HikeKm, 1);
        }

        private static string UnitLabel(ScoutingHistoryInput input, string key)
        {
            var unitIds = input.AllAchievements.results
                .Where(r => Key(r.section) == key && !string.IsNullOrEmpty(r.latest_submission?.unit_id))
                .Select(r => r.latest_submission!.unit_id)
                .Distinct().ToList();

            var known = unitIds.Where(id => input.UnitNames.ContainsKey(id)).Select(id => input.UnitNames[id]).Distinct();
            var unknown = unitIds.Where(id => !input.UnitNames.ContainsKey(id)).ToList();
            var names = known.ToList();
            for (var i = 0; i < unknown.Count; i++)
            {
                // the group's own units can all be looked up, so an unknown one belongs to another group
                names.Add(unknown.Count == 1
                    ? $"{Title(key)} unit (another group)"
                    : $"{Title(key)} unit (another group {i + 1})");
            }
            return string.Join(", ", names);
        }

        private static DateTime? EarliestRecordDate(ScoutingHistoryInput input, string key)
        {
            var dates = new List<DateTime>();
            dates.AddRange(input.Milestones.results.Where(r => Key(r.section) == key).Select(r => AwardedDate(r.imported.date_awarded, r.status_updated)));
            dates.AddRange(input.Oas.results.Where(r => Key(r.section) == key).Select(r => AwardedDate(r.imported.date_awarded, r.status_updated)));
            dates.AddRange(input.Sia.results.Where(r => Key(r.section) == key).Select(r => AwardedDate(r.imported.date_awarded, r.status_updated)));
            dates.AddRange(input.AdditionalAwards.results.Where(r => Key(r.section) == key).Select(r => AwardedDate(r.imported.date_awarded, r.status_updated)));
            return dates.Count == 0 ? null : dates.Min();
        }

        // Imported records carry the date they were awarded; otherwise it is when the status last changed
        public static DateTime AwardedDate(string importedDate, DateTime statusUpdated)
        {
            return DateTime.TryParseExact(importedDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
                ? date
                : statusUpdated.Date;
        }

        public static string Duration(DateTime? from, DateTime? to)
        {
            if (from == null || to == null || to < from)
                return "";
            var months = (to.Value.Year - from.Value.Year) * 12 + to.Value.Month - from.Value.Month;
            if (to.Value.Day < from.Value.Day)
                months--;
            if (months < 1)
                return "less than a month";
            var years = months / 12;
            months %= 12;
            var parts = new List<string>();
            if (years > 0)
                parts.Add($"{years} year{(years == 1 ? "" : "s")}");
            if (months > 0)
                parts.Add($"{months} month{(months == 1 ? "" : "s")}");
            return string.Join(" ", parts);
        }

        private static string Key(string? section) => (section ?? "").Trim().ToLowerInvariant();
        private static int SectionRank(string key)
        {
            var rank = Array.IndexOf(SectionOrder, key);
            return rank < 0 ? int.MaxValue : rank;
        }
        private static string SectionName(string key) => Title(key) + "s";
        public static string Title(string text) =>
            CultureInfo.InvariantCulture.TextInfo.ToTitleCase((text ?? "").Replace("_", " ").Replace("-", " ").Trim());
        private static string StatusText(string status) => status == "awarded" ? "Awarded" : status == "in_progress" ? "In progress" : Title(status);
    }
}
