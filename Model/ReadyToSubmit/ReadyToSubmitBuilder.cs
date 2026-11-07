using System.Globalization;
using Topo.Model.Members;

namespace Topo.Model.ReadyToSubmit
{
    public class ReadyToSubmitInput
    {
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public DateTime GeneratedAt { get; set; } = DateTime.Now;
        // In the order they are to be listed
        public List<ReadyMemberInput> Members { get; set; } = new List<ReadyMemberInput>();
        public ReadyTemplates Templates { get; set; } = new ReadyTemplates();
        // An OAS stage ends with "My logbook is up to date", which only the Scout can tick. When this is set a stage that has
        // everything else done is listed too, and says what is left.
        public bool LogbookTickOptional { get; set; }
        // Also list OAS stages up to this many ticks short, the commonest missing item first (0 for none)
        public int MaxTicksShort { get; set; }
        // Otherwise a Scout is named by first name, with enough of the surname to tell Scouts apart (as on the sign in sheet)
        public bool UseFullNames { get; set; }
        public List<ReadyCategory> Categories { get; set; } = new List<ReadyCategory>
        {
            ReadyCategory.Milestone, ReadyCategory.OutdoorAdventureSkill, ReadyCategory.SpecialInterestArea, ReadyCategory.PeakAward
        };
    }

    // Which badges have everything done but have not been sent in. Terrain keeps a badge being worked on as a draft (or as in
    // progress) until it is sent for review, so a draft with nothing left to do is the one waiting for a leader to send it.
    public static class ReadyToSubmitBuilder
    {
        public const string Type_Milestone = "milestone";
        public const string Type_Oas = "outdoor_adventure_skill";
        public const string Type_Sia = "special_interest_area";
        public const string LogbookInput = "logbook_up_to_date";

        // The order of the Outdoor Adventure Skills on the wall chart
        private static readonly string[] OasOrder = { "bushcraft", "bushwalking", "camping", "alpine", "cycling", "vertical", "aquatics", "boating", "paddling" };
        private static readonly string[] PeakAwardTypes = { "adventurous_journey", "course_reflection", "personal_reflection" };
        // Questions that are shown only for some answers, so being blank doesn't mean unfinished
        private static readonly string[] OptionalIds = { "which_goals", "who_is_helping" };
        // Things on a form that are not answered in it
        private static readonly string[] NotQuestions = { "markdown_text", "file_uploader", "image", "divider", "heading", "paragraph" };

        public static ReadyCategory? CategoryOf(string type)
        {
            if (type == Type_Milestone)
                return ReadyCategory.Milestone;
            if (type == Type_Oas)
                return ReadyCategory.OutdoorAdventureSkill;
            if (type == Type_Sia)
                return ReadyCategory.SpecialInterestArea;
            if (PeakAwardTypes.Contains(type))
                return ReadyCategory.PeakAward;
            return null;
        }

        // The template that tells whether a form is finished: a plan waiting for approval is judged by its plan, a review by its review
        public static string FormTemplatePath(AchievementRecord record)
        {
            if (record.Status == "draft_review" && record.Template.EndsWith("/plan", StringComparison.OrdinalIgnoreCase))
                return record.Template.Substring(0, record.Template.Length - "plan".Length) + "review";
            return record.Template;
        }

        // The records that could be waiting and so need their template read to find out
        public static bool NeedsFormTemplate(AchievementRecord record, string section) =>
            IsInSection(record, section) && IsFormType(record.Type) && (record.Status == "draft_approval" || record.Status == "draft_review") && record.Template != "";

        public static bool NeedsOasTemplate(AchievementRecord record, string section) =>
            IsInSection(record, section) && record.Type == Type_Oas && record.Status == "draft_review" && record.Template != "" && record.Answers.Count > 0;

        public static ReadyToSubmitReport Build(ReadyToSubmitInput input)
        {
            var section = (input.Section ?? "").Trim().ToLowerInvariant();
            var report = new ReadyToSubmitReport
            {
                UnitName = input.UnitName,
                Section = section,
                GeneratedAt = input.GeneratedAt,
                Categories = input.Categories.ToList(),
                MembersChecked = input.Members.Count,
                MaxTicksShort = Math.Max(0, input.MaxTicksShort),
                LogbookTickOptional = input.LogbookTickOptional
            };

            var names = DisplayNames(input.Members, input.UseFullNames);
            var ready = new List<ReadyItem>();
            var logbook = new List<ReadyItem>();
            var nearly = new List<(ReadyItem item, int order)>();
            var order = 0;
            foreach (var member in input.Members)
            {
                order++;
                var items = new List<ReadyItem>();
                foreach (var record in member.Records.Where(r => IsInSection(r, section)))
                {
                    var category = CategoryOf(record.Type);
                    if (category == null || !input.Categories.Contains(category.Value))
                        continue;
                    var verdict = Judge(record, input);
                    if (verdict == null)
                        continue;
                    items.Add(new ReadyItem
                    {
                        Kind = verdict.Kind,
                        MemberId = member.MemberId,
                        MemberName = names[member],
                        Patrol = member.Patrol,
                        Category = category.Value,
                        Badge = BadgeName(record),
                        BadgeOrder = BadgeOrderOf(record),
                        Waiting = verdict.Waiting,
                        Missing = verdict.Missing,
                        MissingPositions = verdict.MissingPositions,
                        MemberOrder = order,
                        LastUpdated = record.StatusUpdated
                    });
                }
                foreach (var item in items.OrderBy(i => i.Category).ThenBy(i => i.Badge, StringComparer.OrdinalIgnoreCase))
                {
                    if (item.Kind == ReadyKind.Ready)
                        ready.Add(item);
                    else if (item.Kind == ReadyKind.LogbookTickOnly)
                        logbook.Add(item);
                    else
                        nearly.Add((item, order));
                }
            }

            report.Items.AddRange(ready);
            report.Items.AddRange(logbook);
            report.Items.AddRange(SortByCommonestMissing(nearly));
            return report;
        }

        // The commonest missing item first, so what the most Scouts are stuck on comes at the top and the stages missing the same
        // thing sit together. A stage missing several things goes with the one that most Scouts are missing.
        private static IEnumerable<ReadyItem> SortByCommonestMissing(List<(ReadyItem item, int order)> nearly)
        {
            var count = new Dictionary<string, int>();
            string Key(ReadyItem item, string label) => item.Badge + "|" + label;
            foreach (var (item, _) in nearly)
                foreach (var label in item.Missing)
                    count[Key(item, label)] = count.GetValueOrDefault(Key(item, label)) + 1;

            (string label, int frequency) Commonest(ReadyItem item) =>
                item.Missing.Select(label => (label, frequency: count[Key(item, label)])).OrderByDescending(m => m.frequency).First();

            return nearly
                .OrderByDescending(n => Commonest(n.item).frequency)
                .ThenBy(n => n.item.Badge, StringComparer.OrdinalIgnoreCase)
                .ThenBy(n => Commonest(n.item).label, StringComparer.OrdinalIgnoreCase)
                .ThenBy(n => n.item.TicksShort)
                .ThenBy(n => n.order)
                .Select(n => n.item);
        }

        private class Verdict
        {
            public ReadyKind Kind { get; set; }
            public string Waiting { get; set; } = string.Empty;
            public List<string> Missing { get; set; } = new List<string>();
            public List<int> MissingPositions { get; set; } = new List<int>();
        }

        // Where the record belongs, or null when it isn't ready (or near enough)
        private static Verdict? Judge(AchievementRecord record, ReadyToSubmitInput input)
        {
            var templates = input.Templates;
            if (record.Type == Type_Milestone)
                return (record.Status == "in_progress" || record.Status == "draft_review") && record.RequirementStatus == "complete"
                    ? new Verdict { Kind = ReadyKind.Ready, Waiting = "Requirements complete, not sent for review" }
                    : null;

            if (record.Type == Type_Oas)
            {
                if (record.Status != "draft_review" || !templates.OasInputs.TryGetValue(record.Template, out var inputs) || inputs.Count == 0)
                    return null;
                var left = inputs.Where(q => !OasInputDone(record, q.Id)).ToList();
                var positions = left.Select(q => inputs.IndexOf(q)).ToList();
                if (left.Count == 0)
                    return new Verdict { Kind = ReadyKind.Ready, Waiting = "Every requirement ticked off, not sent for review" };
                var missing = left.Select(q => q.Label == "" ? q.Id : q.Label).ToList();
                // only the Scout's own logbook tick is left
                if (input.LogbookTickOptional && left.Count == 1 && left[0].Id == LogbookInput)
                    return new Verdict { Kind = ReadyKind.LogbookTickOnly, Waiting = "Everything ticked off but the Scout's logbook tick", Missing = missing, MissingPositions = positions };
                if (left.Count <= input.MaxTicksShort)
                    return new Verdict { Kind = ReadyKind.NearlyThere, Waiting = $"{left.Count} {(left.Count == 1 ? "tick" : "ticks")} short", Missing = missing, MissingPositions = positions };
                return null;
            }

            if (IsFormType(record.Type) && (record.Status == "draft_approval" || record.Status == "draft_review"))
            {
                if (!templates.Forms.TryGetValue(FormTemplatePath(record), out var questions))
                    return null;
                var toAnswer = questions.Where(q => !NotQuestions.Contains(q.Type) && !OptionalIds.Contains(q.Id) && !q.Label.Contains("optional", StringComparison.OrdinalIgnoreCase)).ToList();
                if (toAnswer.Count == 0 || !toAnswer.All(q => Answered(record, q)))
                    return null;
                return new Verdict { Kind = ReadyKind.Ready, Waiting = record.Status == "draft_approval" ? "Plan finished, not sent for approval" : "Finished, not sent for review" };
            }
            return null;
        }

        private static bool OasInputDone(AchievementRecord record, string id)
        {
            if (record.Answers.TryGetValue(id + "_verifiedDate", out var date) && !string.IsNullOrWhiteSpace(date))
                return true;
            // the logbook tick is stored under its own name
            return id == LogbookInput && record.Answers.TryGetValue(id, out var ticked) && ticked == "true";
        }

        private static bool Answered(AchievementRecord record, FormQuestion question)
        {
            if (!record.Answers.TryGetValue(question.Id, out var answer) || string.IsNullOrWhiteSpace(answer))
                return false;
            return question.Type != "checkbox" || answer == "true";
        }

        private static bool IsFormType(string type) => type == Type_Sia || PeakAwardTypes.Contains(type);

        private static bool IsInSection(AchievementRecord record, string section) =>
            string.Equals(record.Section, section, StringComparison.OrdinalIgnoreCase);

        // An Outdoor Adventure Skill stage by its place on the wall chart, then its level; anything else keeps its name's order
        private static int BadgeOrderOf(AchievementRecord record)
        {
            if (record.Type != Type_Oas)
                return int.MaxValue;
            var stream = OasOrder.ToList().IndexOf((record.Meta.TryGetValue("stream", out var s) ? s : "").ToLowerInvariant());
            var stage = record.Meta.TryGetValue("stage", out var l) && int.TryParse(l, out var level) ? level : 0;
            return (stream < 0 ? OasOrder.Length : stream) * 100 + stage;
        }

        // First name only unless that is shared; where several Scouts share a first name the first letters of their surname are
        // added (one more letter at a time) until each is different. Anyone still identical gets a number. The same as the sign in sheet.
        public static Dictionary<ReadyMemberInput, string> DisplayNames(IEnumerable<ReadyMemberInput> members, bool useFullNames)
        {
            var names = new Dictionary<ReadyMemberInput, string>();
            string First(ReadyMemberInput m) => (m.FirstName == "" && m.LastName == "" ? m.MemberName : m.FirstName).Trim();
            string Left(string text, int letters) => text.Length <= letters ? text : text.Substring(0, letters);

            void Name(IEnumerable<ReadyMemberInput> people, Func<ReadyMemberInput, string> name)
            {
                foreach (var same in people.GroupBy(name, StringComparer.OrdinalIgnoreCase))
                {
                    var duplicates = same.ToList();
                    for (var i = 0; i < duplicates.Count; i++)
                        names[duplicates[i]] = duplicates.Count > 1 ? $"{same.Key} ({i + 1})" : same.Key;
                }
            }

            if (useFullNames)
            {
                Name(members, m => $"{First(m)} {m.LastName}".Trim());
                return names;
            }
            foreach (var sameFirstName in members.GroupBy(First, StringComparer.OrdinalIgnoreCase))
            {
                var people = sameFirstName.ToList();
                var letters = 1;
                var longest = people.Max(p => p.LastName.Length);
                while (people.Count > 1 && letters < longest
                    && people.GroupBy(p => Left(p.LastName, letters), StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                    letters++;
                Name(people, p => people.Count == 1 ? First(p) : $"{First(p)} {Left(p.LastName, letters)}".Trim());
            }
            return names;
        }

        private static string BadgeName(AchievementRecord record)
        {
            string Meta(string key) => record.Meta.TryGetValue(key, out var value) ? value : "";
            var text = CultureInfo.InvariantCulture.TextInfo;
            switch (record.Type)
            {
                case Type_Milestone:
                    return $"Milestone {Meta("stage")}";
                case Type_Oas:
                    return $"{text.ToTitleCase(Meta("stream").Replace("_", " "))} stage {Meta("stage")}";
                case Type_Sia:
                    return "SIA: " + (record.Answers.TryGetValue("project_name", out var name) && name != "" ? name : "(unnamed project)");
                case "adventurous_journey":
                    return "Adventurous Journey";
                case "course_reflection":
                    return "Personal Development Course";
                case "personal_reflection":
                    return "Personal Reflection";
                default:
                    return record.Type;
            }
        }

        // The unit leader first, then by patrol (people without one last); within a patrol the patrol leader, then the assistant
        // patrol leader, then everyone else by name. The same order as the sign in sheet.
        public static List<MemberListModel> InPatrolOrder(IEnumerable<MemberListModel> members) =>
            members
                .OrderBy(m => m.patrol_duty == "UL" ? 0 : 1)
                .ThenBy(m => string.IsNullOrWhiteSpace(m.patrol_name) ? 1 : 0)
                .ThenBy(m => m.patrol_name ?? "", StringComparer.OrdinalIgnoreCase)
                .ThenBy(m => m.patrol_duty == "PL" ? 0 : m.patrol_duty == "APL" ? 1 : 2)
                .ThenBy(m => m.first_name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(m => m.last_name, StringComparer.OrdinalIgnoreCase)
                .ToList();
    }
}
