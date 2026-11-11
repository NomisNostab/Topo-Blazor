using System.Globalization;
using System.Text.Json;

namespace Topo.Model.Program
{
    // An activity a youth member has proposed in Terrain (the unit's "Proposed activities"), waiting for the leaders to plan it
    public class ProposedActivityModel
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Justification { get; set; } = string.Empty;
        // Who they suggest should lead it, in their own words
        public string SuggestedLead { get; set; } = string.Empty;
        // Terrain's own name for the area, like "personal_growth"
        public string ChallengeAreaKey { get; set; } = string.Empty;
        public string ChallengeArea => ProposedActivities.AreaLabel(ChallengeAreaKey);
        // Tells one proposal from another on the page, even if Terrain gave it no id
        public string OpenKey => Id != "" ? Id : Title + "|" + Submitted?.Ticks;
        public DateTime? Submitted { get; set; }
        public string ProposedBy { get; set; } = string.Empty;
        public string Patrol { get; set; } = string.Empty;
        // How many times an activity of this name has been proposed (this one included)
        public int TimesProposed { get; set; } = 1;
        // How many times an activity of this name has been run before, and when it last was; null when the unit's events could not be read
        public int? TimesRun { get; set; }
        public DateTime? LastRun { get; set; }
        // The names of the events counted as runs, with how many times each was run, like "Origami Night (2)"
        public List<string> RunTitles { get; set; } = new List<string>();
    }

    // What the Proposed Activities report is made from: the proposals as they are shown, and the date the run counts start from
    public class ProposedActivitiesReportData
    {
        // Events before this date are not counted as runs; null counts all of them
        public DateTime? Since { get; set; }
        public List<ProposedActivityModel> Proposals { get; set; } = new List<ProposedActivityModel>();
    }

    public enum ProposedActivitySortColumn
    {
        Submitted,
        Activity,
        ChallengeArea,
        ProposedBy,
        TimesProposed,
        TimesRun
    }

    public static class ProposedActivities
    {
        public static string AreaLabel(string key)
        {
            return (key ?? "").Trim().ToLowerInvariant() switch
            {
                "outdoors" => "Outdoors",
                "creative" => "Creative",
                "community" => "Community",
                "personal_growth" => "Personal Growth",
                "" or "not_applicable" => "Not applicable",
                var other => other
            };
        }

        // Terrain answers with the proposals either as a list or inside an object holding one; null when there is nothing readable
        public static List<ProposedActivityModel>? Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return null;
            try
            {
                using var document = JsonDocument.Parse(json);
                var list = FindList(document.RootElement);
                if (list == null)
                    return null;

                var result = new List<ProposedActivityModel>();
                foreach (var item in list.Value.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.Object)
                        continue;
                    var title = Text(item, "title").Trim();
                    if (title == "")
                        continue;
                    var member = item.TryGetProperty("proposing_member", out var m) && m.ValueKind == JsonValueKind.Object ? m : (JsonElement?)null;
                    result.Add(new ProposedActivityModel
                    {
                        Id = Text(item, "id"),
                        Title = title,
                        Description = Text(item, "description").Trim(),
                        Justification = Text(item, "justification").Trim(),
                        SuggestedLead = Text(item, "leader").Trim(),
                        ChallengeAreaKey = Text(item, "challenge_area"),
                        Submitted = Date(Text(item, "created_at")),
                        ProposedBy = member == null ? "" : (Text(member.Value, "first_name") + " " + Text(member.Value, "last_name")).Trim(),
                        Patrol = member == null ? "" : Text(member.Value, "patrol_name").Trim()
                    });
                }
                return result;
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static JsonElement? FindList(JsonElement root)
        {
            if (root.ValueKind == JsonValueKind.Array)
                return root;
            if (root.ValueKind != JsonValueKind.Object)
                return null;
            foreach (var name in new[] { "results", "proposals", "data" })
                if (root.TryGetProperty(name, out var named) && named.ValueKind == JsonValueKind.Array)
                    return named;
            foreach (var property in root.EnumerateObject())
                if (property.Value.ValueKind == JsonValueKind.Array)
                    return property.Value;
            return null;
        }

        private static string Text(JsonElement element, string name)
        {
            return element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : "";
        }

        private static DateTime? Date(string text)
        {
            if (DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
                return parsed.LocalDateTime;
            return null;
        }

        // Words that do not tell one activity from another, so "Origami" and "Origami Night" are the same activity
        private static readonly HashSet<string> FillerWords = new HashSet<string>
        {
            "a", "an", "the", "and", "of", "in", "at", "to", "for", "with", "on", "night", "nights", "day", "evening", "activity", "activities", "session"
        };

        // The main words of an activity's name: lower case, no punctuation, no filler, and plurals made singular
        public static HashSet<string> NameWords(string name)
        {
            var words = new List<string>();
            var current = new System.Text.StringBuilder();
            foreach (var character in (name ?? "").ToLowerInvariant() + " ")
            {
                if (char.IsLetterOrDigit(character))
                    current.Append(character);
                else if (current.Length > 0)
                {
                    words.Add(Singular(current.ToString()));
                    current.Clear();
                }
            }
            var main = words.Where(w => !FillerWords.Contains(w)).ToList();
            // a name of nothing but filler ("Night") is still compared by what it has
            return new HashSet<string>(main.Count > 0 ? main : words);
        }

        private static string Singular(string word)
        {
            return word.Length > 3 && word.EndsWith("s") && !word.EndsWith("ss") ? word.Substring(0, word.Length - 1) : word;
        }

        // Two names are the same activity when one has all the main words of the other
        public static bool SameActivity(HashSet<string> first, HashSet<string> second)
        {
            if (first.Count == 0 || second.Count == 0)
                return false;
            return first.IsSubsetOf(second) || second.IsSubsetOf(first);
        }

        // Fills in how often each activity has been proposed, and from the unit's events how often it has been run before now.
        // With no events (they could not be read) the times run are left unknown rather than shown as none.
        public static void AddHistory(List<ProposedActivityModel> proposals, IEnumerable<(string Title, DateTime Start)>? events, DateTime now)
        {
            var proposalWords = proposals.Select(p => NameWords(p.Title)).ToList();
            var past = events?.Where(e => e.Start <= now).Select(e => (e.Title, e.Start, Words: NameWords(e.Title))).ToList();

            for (int i = 0; i < proposals.Count; i++)
            {
                var proposal = proposals[i];
                proposal.TimesProposed = proposalWords.Count(words => SameActivity(proposalWords[i], words));
                proposal.RunTitles = new List<string>();
                if (past == null)
                {
                    proposal.TimesRun = null;
                    proposal.LastRun = null;
                    continue;
                }

                var runs = past.Where(e => SameActivity(proposalWords[i], e.Words)).ToList();
                proposal.TimesRun = runs.Count;
                proposal.LastRun = runs.Count == 0 ? (DateTime?)null : runs.Max(e => e.Start);
                proposal.RunTitles = runs.GroupBy(e => e.Title.Trim(), StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(g => g.Count()).ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.Count() == 1 ? g.Key : $"{g.Key} ({g.Count()})").ToList();
            }
        }

        public static List<ProposedActivityModel> Sort(IEnumerable<ProposedActivityModel> proposals, ProposedActivitySortColumn column, bool descending)
        {
            IOrderedEnumerable<ProposedActivityModel> ordered = column switch
            {
                ProposedActivitySortColumn.Activity => proposals.OrderBy(p => p.Title, StringComparer.OrdinalIgnoreCase),
                ProposedActivitySortColumn.ChallengeArea => proposals.OrderBy(p => p.ChallengeArea, StringComparer.OrdinalIgnoreCase),
                ProposedActivitySortColumn.ProposedBy => proposals.OrderBy(p => p.ProposedBy, StringComparer.OrdinalIgnoreCase),
                ProposedActivitySortColumn.TimesProposed => proposals.OrderBy(p => p.TimesProposed),
                ProposedActivitySortColumn.TimesRun => proposals.OrderBy(p => p.TimesRun ?? -1),
                _ => proposals.OrderBy(p => p.Submitted ?? DateTime.MaxValue)
            };
            ordered = ordered.ThenBy(p => p.Title, StringComparer.OrdinalIgnoreCase);
            var list = ordered.ToList();
            if (descending)
                list.Reverse();
            return list;
        }
    }
}
