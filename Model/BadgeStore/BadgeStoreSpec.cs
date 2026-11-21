using Topo.Model.Program;

namespace Topo.Model.BadgeStore
{
    // Terrain has nowhere to record that an additional award (Camper, Walkabout, First Aid and so on) was awarded or
    // presented, and no call a leader can reach creates one. Topo keeps its own records as comments on one planned group
    // event far in the future. Every Topo install uses the same event title, date and record format below, so any leader in
    // the group, using any copy of Topo, finds and reads the same records. See docs/badge-store.md for the full description.
    public static class BadgeStoreSpec
    {
        // Version of the record format. A record with a higher version is left alone (and counted) rather than guessed at.
        public const int SchemaVersion = 1;

        // The event is found by this start of its title, so a leader can reword the end of it without losing the store.
        public const string TitlePrefix = "Topo records";
        public const string Title = "Topo records - please do not delete";
        public const string Description =
            "Topo records v1. The Topo app (github.com/NomisNostab/Topo-Blazor) keeps a note here, as comments, of badges " +
            "that have been awarded or presented, because Terrain has nowhere to record them. Please do not edit or delete this event.";

        // The event sits on one fixed day, so finding it is a read of one day's calendar
        public static readonly DateTime StartUtc = new DateTime(2100, 2, 1, 8, 0, 0, DateTimeKind.Utc);
        public static readonly DateTime EndUtc = new DateTime(2100, 2, 1, 9, 0, 0, DateTimeKind.Utc);
        // A little either side of the day, in case a client sends the time in another time zone
        public static readonly DateTime SearchFromUtc = new DateTime(2100, 1, 31, 0, 0, 0, DateTimeKind.Utc);
        public static readonly DateTime SearchToUtc = new DateTime(2100, 2, 3, 0, 0, 0, DateTimeKind.Utc);

        // True for the group's own store event in a calendar read
        public static bool IsStoreEvent(Event e, string groupId) =>
            e.invitee_type == "group"
            && e.invitee_id == groupId
            && e.title.StartsWith(TitlePrefix, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(e.status, "cancelled", StringComparison.OrdinalIgnoreCase);
    }
}
