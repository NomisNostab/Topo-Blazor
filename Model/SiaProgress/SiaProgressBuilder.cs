using System.Globalization;
using Topo.Model.Members;
using Topo.Model.SIA;

namespace Topo.Model.SiaProgress
{
    public class SiaMemberInput
    {
        public string MemberId { get; set; } = string.Empty;
        public string MemberName { get; set; } = string.Empty;
        public GetSIAResultsModel Results { get; set; } = new GetSIAResultsModel();
    }

    public class SiaProgressInput
    {
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public List<SiaMemberInput> Members { get; set; } = new List<SiaMemberInput>();
        public SiaTemplates Templates { get; set; } = new SiaTemplates();
        // For a report on one project: its id. Otherwise every project of the members.
        public string ProjectId { get; set; } = string.Empty;
        public bool ExcludeCompleted { get; set; }
        public bool IncludeSummary { get; set; }
        public bool DoubleSided { get; set; }
        public int SiaRequired { get; set; }
        public int AreasRequired { get; set; }
        public string AwardName { get; set; } = string.Empty;
    }

    public static class SiaProgressBuilder
    {
        // The answers that are shown as the project's details rather than as questions
        private static readonly string[] DetailIds = { "special_interest_area_selection", "project_name", "attach_documentation" };
        // Left out when blank: they only apply to some answers (SDGs when connected to them, helpers when not alone)
        private static readonly string[] OptionalIds = { "which_goals", "who_is_helping" };

        public static SiaProgressReport Build(SiaProgressInput input)
        {
            var section = (input.Section ?? "").Trim().ToLowerInvariant();
            var report = new SiaProgressReport
            {
                UnitName = input.UnitName,
                Section = section,
                IncludeSummary = input.IncludeSummary,
                DoubleSided = input.DoubleSided,
                AwardName = input.AwardName,
                SiaRequired = input.SiaRequired,
                AreasRequired = input.AreasRequired
            };

            // in the order the members are given
            foreach (var member in input.Members)
            {
                // only this section's projects count, as they do on the Wallchart
                var projects = member.Results.results.Where(r => Key(r.section) == section).ToList();

                if (input.IncludeSummary)
                {
                    var complete = projects.Where(IsComplete).ToList();
                    report.Summary.Add(new SiaSummaryRow
                    {
                        MemberName = member.MemberName,
                        Completed = complete.Count,
                        InProgress = projects.Count - complete.Count,
                        AreasCompleted = complete.Select(r => r.answers.special_interest_area_selection).Where(a => !string.IsNullOrEmpty(a)).Distinct().Count()
                    });
                }

                var shown = projects
                    .Where(r => input.ProjectId == "" || r.id == input.ProjectId)
                    .Where(r => !(input.ExcludeCompleted && input.ProjectId == "" && IsComplete(r)))
                    .OrderBy(r => r.status_updated);
                foreach (var project in shown)
                    report.Projects.Add(ToProject(project, member, input.Templates));
            }
            return report;
        }

        // The projects a Scout has in the section, oldest first, to pick one to report on
        public static List<SiaProjectChoice> ProjectChoices(GetSIAResultsModel results, string section) =>
            results.results
                .Where(r => Key(r.section) == Key(section))
                .OrderBy(r => r.status_updated)
                .Select(r => new SiaProjectChoice
                {
                    Id = r.id,
                    Label = $"{(string.IsNullOrWhiteSpace(r.answers.project_name) ? "(no name yet)" : r.answers.project_name)} - {StatusText(r.status)}"
                })
                .ToList();

        // First name only unless that is shared; where several Scouts share a first name the first letters of their surname are added
        // (one more letter at a time) until each is different. Anyone still identical gets a number. The same as the sign in sheet.
        public static Dictionary<string, string> ShortNames(IEnumerable<MemberListModel> members)
        {
            var names = new Dictionary<string, string>();
            string Left(string? value, int length) { value ??= ""; return value.Length <= length ? value : value.Substring(0, length); }

            foreach (var sameFirstName in members.GroupBy(m => (m.first_name ?? "").Trim(), StringComparer.OrdinalIgnoreCase))
            {
                var people = sameFirstName.ToList();
                var first = people[0].first_name?.Trim() ?? "";
                if (people.Count == 1)
                {
                    names[people[0].id] = first;
                    continue;
                }

                var longest = people.Max(p => (p.last_name ?? "").Length);
                var letters = 1;
                while (letters < longest && people.GroupBy(p => Left(p.last_name, letters), StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                    letters++;

                foreach (var same in people.GroupBy(p => $"{first} {Left(p.last_name, letters)}".Trim(), StringComparer.OrdinalIgnoreCase))
                {
                    var duplicates = same.ToList();
                    for (var i = 0; i < duplicates.Count; i++)
                        names[duplicates[i].id] = duplicates.Count > 1 ? $"{same.Key} ({i + 1})" : same.Key;
                }
            }
            return names;
        }

        // The unit leader first, then by patrol (people without one last); within a patrol the patrol leader, then the
        // assistant patrol leader, then everyone else by name. The same order as the sign in sheet.
        public static List<MemberListModel> InPatrolOrder(IEnumerable<MemberListModel> members) =>
            members
                .OrderBy(m => m.patrol_duty == "UL" ? 0 : 1)
                .ThenBy(m => string.IsNullOrWhiteSpace(m.patrol_name) ? 1 : 0)
                .ThenBy(m => m.patrol_name ?? "", StringComparer.OrdinalIgnoreCase)
                .ThenBy(m => m.patrol_duty == "PL" ? 0 : m.patrol_duty == "APL" ? 1 : 2)
                .ThenBy(m => m.first_name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(m => m.last_name, StringComparer.OrdinalIgnoreCase)
                .ToList();

        private static bool IsComplete(Result project) => project.status == "awarded";

        private static SiaProject ToProject(Result r, SiaMemberInput member, SiaTemplates templates)
        {
            DateTime? awarded = null;
            if (IsComplete(r))
                awarded = DateTime.TryParseExact(r.imported?.date_awarded, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var imported)
                    ? imported
                    : r.status_updated.Date;

            var options = templates.Plan.Options.Concat(templates.Review.Options)
                .GroupBy(o => o.Key).ToDictionary(g => g.Key, g => g.First().Value);
            return new SiaProject
            {
                MemberId = member.MemberId,
                MemberName = member.MemberName,
                ProjectName = r.answers.project_name,
                Area = AreaName(r.answers.special_interest_area_selection, options),
                Status = StatusText(r.status),
                IsComplete = IsComplete(r),
                StatusUpdated = r.status_updated,
                DateAwarded = awarded,
                Plan = Answers(r.answers, templates.Plan, options),
                Review = Answers(r.answers, templates.Review, options),
                Files = r.uploads.Select(u => u.filename).ToList()
            };
        }

        private static List<SiaAnswer> Answers(Answers answers, SiaTemplate template, Dictionary<string, string> options)
        {
            var list = new List<SiaAnswer>();
            // a question this report has no answer field for would only ever read "not answered", so it is left out
            foreach (var question in template.Questions.Where(q => !DetailIds.Contains(q.Id) && typeof(Answers).GetProperty(q.Id) != null))
            {
                var text = AnswerText(answers, question.Id, options);
                if (text == "" && OptionalIds.Contains(question.Id))
                    continue;
                list.Add(new SiaAnswer { Question = question.Label, Answer = text });
            }
            return list;
        }

        // What the Scout answered, by Terrain's name for the question, with a chosen option shown as its wording
        private static string AnswerText(Answers answers, string id, Dictionary<string, string> options)
        {
            var value = typeof(Answers).GetProperty(id)?.GetValue(answers);
            var values = value is string[] many ? many : new[] { value as string ?? "" };
            return string.Join(", ", values.Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => options.TryGetValue(v, out var title) ? title : v.Trim()));
        }

        private static string AreaName(string key, Dictionary<string, string> options)
        {
            if (string.IsNullOrEmpty(key))
                return "";
            return options.TryGetValue(key, out var title) ? title : CultureInfo.InvariantCulture.TextInfo.ToTitleCase(key.Replace("sia_", "").Replace("_", " "));
        }

        private static string StatusText(string status)
        {
            var text = (status ?? "").Replace("_", " ");
            return text.Length == 0 ? "" : char.ToUpperInvariant(text[0]) + text.Substring(1);
        }

        private static string Key(string? section) => (section ?? "").Trim().ToLowerInvariant();
    }
}
