using System.Text.Json;
using Topo.Model.Program;

namespace Topo.Model.BadgeStore
{
    // Terrain's calendar read returns only events from calendars that are switched on, and a leader's group calendar is
    // usually off. To find the store event Topo switches the group calendar on, reads one day, and puts the leader's own
    // selection back exactly as it was.
    public static class BadgeStoreCalendars
    {
        // The selection to send to read the store, or null when no change is needed: the group calendar is already on, or
        // the leader has no calendar for the group. The calendars passed in are not changed, so they can be sent back as they were.
        public static GetCalendarsResultModel? WithGroupCalendarOn(GetCalendarsResultModel original, string groupId)
        {
            var group = original.own_calendars.FirstOrDefault(c => c.type == "group" && c.id == groupId)
                        ?? original.own_calendars.FirstOrDefault(c => c.id == groupId);
            if (group == null || group.selected)
                return null;

            var copy = JsonSerializer.Deserialize<GetCalendarsResultModel>(JsonSerializer.Serialize(original))!;
            var copyGroup = copy.own_calendars.First(c => c.id == group.id);
            copyGroup.selected = true;
            return copy;
        }
    }
}
