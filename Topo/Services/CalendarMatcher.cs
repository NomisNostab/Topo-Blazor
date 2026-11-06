namespace Topo.Services
{
    public static class CalendarMatcher
    {
        // Calendars in a few columns, the shortest names together, so each column need only be as wide as its longest name
        public static List<List<KeyValuePair<string, string>>> Columns(IEnumerable<KeyValuePair<string, string>> calendars)
        {
            var sorted = calendars.OrderBy(c => c.Value.Length).ThenBy(c => c.Value).ToList();
            var rows = Math.Max(3, (int)Math.Ceiling(sorted.Count / 4.0));
            return sorted.Chunk(rows).Select(column => column.OrderBy(c => c.Value).ToList()).ToList();
        }

        // The calendar named after a unit; failing that, one whose name contains the unit's name (or the other way round).
        // Null when there is none.
        public static string? FindForUnit(IEnumerable<KeyValuePair<string, string>> calendars, string? unitName)
        {
            if (string.IsNullOrWhiteSpace(unitName))
                return null;

            var name = unitName.Trim();
            var titled = calendars.Select(c => (id: c.Key, title: c.Value.Trim())).ToList();
            return titled.FirstOrDefault(c => string.Equals(c.title, name, StringComparison.OrdinalIgnoreCase)).id
                ?? titled.FirstOrDefault(c => c.title.Contains(name, StringComparison.OrdinalIgnoreCase)).id
                ?? titled.FirstOrDefault(c => c.title.Length > 0 && name.Contains(c.title, StringComparison.OrdinalIgnoreCase)).id;
        }
    }
}
