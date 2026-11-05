namespace Topo.Model.Members
{
    public class SignInSheetGroup
    {
        // Null when there is nothing to label (all youth are in one section)
        public string? Heading { get; set; }
        public List<MemberListModel> Members { get; set; } = new List<MemberListModel>();
    }

    // How a sign in sheet names and groups people. Shared by the report generator and the on-page
    // preview so what is previewed is what gets printed.
    public static class SignInSheetLayout
    {
        private static readonly string[] SectionOrder = { "joey", "cub", "scout", "venturer", "rover" };

        // First name only by default; where several people share a first name, the first letters of
        // their surname are added (one more letter at a time) until each name is unique. Anyone whose
        // name is still identical once the whole surname is used gets a number to tell them apart.
        public static Dictionary<string, string> GetDisplayNames(IEnumerable<MemberListModel> members, bool useFullNames)
        {
            var displayNames = new Dictionary<string, string>();
            if (useFullNames)
            {
                foreach (var member in members)
                    displayNames[member.id] = $"{member.first_name} {member.last_name}".Trim();
                return displayNames;
            }

            foreach (var sameFirstName in members.GroupBy(m => (m.first_name ?? "").Trim(), StringComparer.OrdinalIgnoreCase))
            {
                var people = sameFirstName.ToList();
                var firstName = people[0].first_name?.Trim() ?? "";
                if (people.Count == 1)
                {
                    displayNames[people[0].id] = firstName;
                    continue;
                }

                var longestLastName = people.Max(p => (p.last_name ?? "").Length);
                var letters = 1;
                while (letters < longestLastName
                    && people.GroupBy(p => Left(p.last_name, letters), StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                {
                    letters++;
                }

                foreach (var sameName in people.GroupBy(p => $"{firstName} {Left(p.last_name, letters)}".Trim(), StringComparer.OrdinalIgnoreCase))
                {
                    var duplicates = sameName.ToList();
                    for (var i = 0; i < duplicates.Count; i++)
                        displayNames[duplicates[i].id] = duplicates.Count > 1 ? $"{sameName.Key} ({i + 1})" : sameName.Key;
                }
            }
            return displayNames;
        }

        // Youth grouped by section (Joeys first through Rovers) when members came from more than one
        // section, otherwise one list; adult leaders are always listed separately at the end.
        public static List<SignInSheetGroup> GetGroups(IEnumerable<MemberListModel> members, Dictionary<string, string> displayNames)
        {
            var memberList = members.ToList();
            // The unit leader first, then by patrol (people without one last); within a patrol the patrol
            // leader then the assistant patrol leader then everyone else alphabetically by the name shown.
            List<MemberListModel> Ordered(IEnumerable<MemberListModel> people) =>
                people
                    .OrderBy(m => m.patrol_duty == "UL" ? 0 : 1)
                    .ThenBy(m => string.IsNullOrWhiteSpace(m.patrol_name) ? 1 : 0)
                    .ThenBy(m => m.patrol_name ?? "", StringComparer.OrdinalIgnoreCase)
                    .ThenBy(m => DutyRank(m.patrol_duty))
                    .ThenBy(m => displayNames[m.id], StringComparer.OrdinalIgnoreCase)
                    .ToList();

            var groups = new List<SignInSheetGroup>();
            var youthSections = memberList
                .Where(m => m.isAdultLeader == 0)
                .GroupBy(m => (m.section ?? "").ToLowerInvariant())
                .OrderBy(g => Array.IndexOf(SectionOrder, g.Key) < 0 ? int.MaxValue : Array.IndexOf(SectionOrder, g.Key))
                .ThenBy(g => g.Key)
                .ToList();

            foreach (var youthSection in youthSections)
            {
                var heading = youthSections.Count <= 1 ? null
                    : string.IsNullOrEmpty(youthSection.Key)
                        ? "Other"
                        : System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase(youthSection.Key) + "s";
                groups.Add(new SignInSheetGroup { Heading = heading, Members = Ordered(youthSection) });
            }

            var leaders = memberList.Where(m => m.isAdultLeader == 1).ToList();
            if (leaders.Any())
                groups.Add(new SignInSheetGroup { Heading = "Leaders", Members = Ordered(leaders) });

            return groups;
        }

        // Patrol leader (a Cub "Sixer" counts as one) before assistant patrol leader (a "Seconder"), then everyone else.
        private static int DutyRank(string? duty)
        {
            var name = (duty ?? "").Trim();
            if (name.Equals("PL", StringComparison.OrdinalIgnoreCase) || name.Equals("Sixer", StringComparison.OrdinalIgnoreCase)
                || name.Equals("Patrol Leader", StringComparison.OrdinalIgnoreCase))
                return 0;
            if (name.Equals("APL", StringComparison.OrdinalIgnoreCase) || name.Equals("Seconder", StringComparison.OrdinalIgnoreCase)
                || name.Equals("Assistant Patrol Leader", StringComparison.OrdinalIgnoreCase))
                return 1;
            return 2;
        }

        private static string Left(string? value, int length)
        {
            value ??= "";
            return value.Length <= length ? value : value.Substring(0, length);
        }
    }
}
