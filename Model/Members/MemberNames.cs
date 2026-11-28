namespace Topo.Model.Members
{
    // How people are named on a report, so every report that names people names them the same way.
    public static class MemberNames
    {
        // First name only by default; where several people share a first name, the first letters of their
        // surname are added (one more letter at a time) until each name is unique. Anyone whose name is
        // still identical once the whole surname is used gets a number to tell them apart. With useFullNames
        // it is first name and surname, numbered the same way.
        // Works on any kind of person: say where each one's first name and surname come from.
        public static Dictionary<T, string> GetDisplayNames<T>(IEnumerable<T> people, Func<T, string?> firstName, Func<T, string?> lastName, bool useFullNames)
            where T : notnull
        {
            var names = new Dictionary<T, string>();
            string First(T person) => (firstName(person) ?? "").Trim();
            string Last(T person) => (lastName(person) ?? "").Trim();

            void NumberDuplicates(IEnumerable<T> group, Func<T, string> name)
            {
                foreach (var same in group.GroupBy(name, StringComparer.OrdinalIgnoreCase))
                {
                    var duplicates = same.ToList();
                    for (var i = 0; i < duplicates.Count; i++)
                        names[duplicates[i]] = duplicates.Count > 1 ? $"{same.Key} ({i + 1})" : same.Key;
                }
            }

            if (useFullNames)
            {
                NumberDuplicates(people, p => $"{First(p)} {Last(p)}".Trim());
                return names;
            }

            foreach (var sameFirstName in people.GroupBy(First, StringComparer.OrdinalIgnoreCase))
            {
                var sharing = sameFirstName.ToList();
                var first = First(sharing[0]);
                if (sharing.Count == 1)
                {
                    names[sharing[0]] = first;
                    continue;
                }

                var longestLastName = sharing.Max(p => Last(p).Length);
                var letters = 1;
                while (letters < longestLastName
                    && sharing.GroupBy(p => Left(Last(p), letters), StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                {
                    letters++;
                }
                NumberDuplicates(sharing, p => $"{first} {Left(Last(p), letters)}".Trim());
            }
            return names;
        }

        private static string Left(string value, int length) => value.Length <= length ? value : value.Substring(0, length);
    }
}
