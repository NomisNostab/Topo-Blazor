using Topo.Model.BadgeStore;
using Topo.Model.Program;

namespace Topo.Services
{
    public interface IBadgeStoreService
    {
        // The group's store events (normally one), found with the group calendar switched on and put back afterwards
        Task<List<Event>> FindStoreEventsAsync(string userId, string groupId);
        // Every record in the group's store, from comments written by someone isTrustedAuthor accepts
        Task<BadgeStoreContents> ReadAsync(string userId, string groupId, Func<string, bool> isTrustedAuthor);
    }

    // Reads the badge records Topo keeps in Terrain (Model/BadgeStore/BadgeStoreSpec.cs, docs/badge-store.md). Read only.
    public class BadgeStoreService : IBadgeStoreService
    {
        // The calendar selection belongs to the account, so one read at a time switches it and puts it back
        private static readonly SemaphoreSlim CalendarGate = new SemaphoreSlim(1, 1);

        private readonly ITerrainAPIService _terrainAPIService;

        public BadgeStoreService(ITerrainAPIService terrainAPIService)
        {
            _terrainAPIService = terrainAPIService;
        }

        public async Task<List<Event>> FindStoreEventsAsync(string userId, string groupId)
        {
            await CalendarGate.WaitAsync();
            try
            {
                // Read the selection now, not from an earlier page, so what is put back is what the leader has at this moment
                var original = await _terrainAPIService.GetCalendarsAsync(userId);
                var changed = original == null ? null : BadgeStoreCalendars.WithGroupCalendarOn(original, groupId);
                if (changed != null)
                    await _terrainAPIService.PutCalendarsAsync(userId, changed);
                try
                {
                    var events = await _terrainAPIService.GetEventsAsync(userId, BadgeStoreSpec.SearchFromUtc, BadgeStoreSpec.SearchToUtc);
                    return (events?.results ?? new Event[0]).Where(e => BadgeStoreSpec.IsStoreEvent(e, groupId)).ToList();
                }
                finally
                {
                    if (changed != null)
                        await _terrainAPIService.PutCalendarsAsync(userId, original!);
                }
            }
            finally
            {
                CalendarGate.Release();
            }
        }

        public async Task<BadgeStoreContents> ReadAsync(string userId, string groupId, Func<string, bool> isTrustedAuthor)
        {
            var comments = new List<EventComment>();
            foreach (var storeEvent in await FindStoreEventsAsync(userId, groupId))
                comments.AddRange((await _terrainAPIService.GetEventCommentsAsync(storeEvent.id)).results);
            return BadgeStoreReader.Resolve(comments, isTrustedAuthor);
        }
    }
}
