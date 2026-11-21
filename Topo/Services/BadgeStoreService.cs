using Newtonsoft.Json;
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

        // Creates the group's store event. Only call this once the leader has agreed: it adds an event everyone in the group can see.
        Task<BadgeStoreWriteResult> CreateStoreEventAsync(string userId, string userName, string groupId);
        // Records that a member was awarded (and optionally presented) an additional award
        Task<BadgeStoreWriteResult> RecordAwardAsync(string userId, string groupId, string memberId, string awardId, DateTime awarded, DateTime? presented, Func<string, bool> isTrustedAuthor);
        // Takes a record back
        Task<BadgeStoreWriteResult> RevokeAsync(string userId, string groupId, string memberId, string awardId, Func<string, bool> isTrustedAuthor);
        // Deletes one comment, for undoing a record the signed-in leader wrote
        Task<BadgeStoreWriteResult> DeleteCommentAsync(string eventId, string commentId);
    }

    // Reads and writes the badge records Topo keeps in Terrain (Model/BadgeStore/BadgeStoreSpec.cs, docs/badge-store.md).
    // Everything that writes is asked for by a caller that has already shown the leader what will be written.
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
            var events = await FindStoreEventsAsync(userId, groupId);
            return await ReadEventsAsync(events, isTrustedAuthor);
        }

        private async Task<BadgeStoreContents> ReadEventsAsync(IEnumerable<Event> storeEvents, Func<string, bool> isTrustedAuthor)
        {
            var comments = new List<EventComment>();
            foreach (var storeEvent in storeEvents)
            {
                var read = await _terrainAPIService.GetEventCommentsAsync(storeEvent.id);
                foreach (var comment in read.results)
                    comment.EventId = storeEvent.id;
                comments.AddRange(read.results);
            }
            return BadgeStoreReader.Resolve(comments, isTrustedAuthor);
        }

        public async Task<BadgeStoreWriteResult> CreateStoreEventAsync(string userId, string userName, string groupId)
        {
            var existing = BadgeStoreWrites.ChooseWriteEvent(await FindStoreEventsAsync(userId, groupId));
            if (existing != null)
                return new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.StoreAlreadyExists, EventId = existing.id, Message = "The group already has a store event." };

            var body = JsonConvert.SerializeObject(BadgeStoreWrites.CreateEventBody(groupId, userId, userName));
            var created = await _terrainAPIService.CreateGroupEventAsync(groupId, body);
            if (!created.Success)
                return Failed($"Terrain did not create the store event ({Describe(created)}). Creating a group event needs a group leader.");

            // Terrain does not say which event it made. The organiser's own list shows it whatever calendars are switched on.
            for (var attempt = 0; attempt < 5; attempt++)
            {
                if (attempt > 0)
                    await Task.Delay(1000);
                var organised = await _terrainAPIService.GetOrganisedEventsAsync(userId);
                var found = (organised?.results ?? new Event[0]).Where(e => BadgeStoreWrites.IsCreatedStoreEvent(e, groupId)).OrderBy(e => e.id, StringComparer.Ordinal).FirstOrDefault();
                if (found != null)
                    return new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.StoreCreated, EventId = found.id, Message = "The store event was created." };
            }
            return Failed("Terrain accepted the store event but it could not be found afterwards. Look for it in your organised events before trying again.");
        }

        public async Task<BadgeStoreWriteResult> RecordAwardAsync(string userId, string groupId, string memberId, string awardId, DateTime awarded, DateTime? presented, Func<string, bool> isTrustedAuthor)
        {
            if (string.IsNullOrWhiteSpace(memberId) || string.IsNullOrWhiteSpace(awardId))
                return Failed("A member and an award are needed.");

            var events = await FindStoreEventsAsync(userId, groupId);
            var storeEvent = BadgeStoreWrites.ChooseWriteEvent(events);
            if (storeEvent == null)
                return new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.NoStoreEvent, Message = "The group has no store event yet." };

            var current = await ReadEventsAsync(events, isTrustedAuthor);
            if (BadgeStoreWrites.IsAlreadyRecorded(current, memberId, awardId, awarded, presented))
                return new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.AlreadyRecorded, EventId = storeEvent.id, Message = "The store already says this." };

            return await PostAndConfirmAsync(userId, storeEvent.id, BadgeStoreReader.AwardText(memberId, awardId, awarded, presented));
        }

        public async Task<BadgeStoreWriteResult> RevokeAsync(string userId, string groupId, string memberId, string awardId, Func<string, bool> isTrustedAuthor)
        {
            var events = await FindStoreEventsAsync(userId, groupId);
            var storeEvent = BadgeStoreWrites.ChooseWriteEvent(events);
            if (storeEvent == null)
                return new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.NoStoreEvent, Message = "The group has no store event yet." };

            var current = await ReadEventsAsync(events, isTrustedAuthor);
            if (!BadgeStoreWrites.HasRecord(current, memberId, awardId))
                return new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.NothingToRevoke, EventId = storeEvent.id, Message = "There is no record to take back." };

            return await PostAndConfirmAsync(userId, storeEvent.id, BadgeStoreReader.RevokeText(memberId, awardId));
        }

        public async Task<BadgeStoreWriteResult> DeleteCommentAsync(string eventId, string commentId)
        {
            var deleted = await _terrainAPIService.DeleteEventCommentAsync(eventId, commentId);
            return deleted.Success
                ? new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.Written, EventId = eventId, CommentId = commentId, Message = "The comment was deleted." }
                : Failed($"Terrain did not delete the comment ({Describe(deleted)}).");
        }

        // Posts the comment, then reads the event's comments back to confirm it is there (Terrain returns no id)
        private async Task<BadgeStoreWriteResult> PostAndConfirmAsync(string userId, string eventId, string text)
        {
            var posted = await _terrainAPIService.PostEventCommentAsync(eventId, userId, text);
            if (!posted.Success)
                return Failed($"Terrain did not accept the record ({Describe(posted)}).");

            var read = await _terrainAPIService.GetEventCommentsAsync(eventId);
            var ours = read.results.Where(c => c.text == text && c.member.id == userId).OrderByDescending(c => c.created_datetime).FirstOrDefault();
            if (ours == null)
                return Failed("Terrain accepted the record but it could not be read back. Check the store before trying again.");
            return new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.Written, EventId = eventId, CommentId = ours.id, Message = "The record was saved." };
        }

        private static BadgeStoreWriteResult Failed(string message) => new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.Failed, Message = message };

        private static string Describe(TerrainCallResult result) =>
            result.StatusCode == 0 ? "no reply, which usually means Terrain refused it" : $"HTTP {result.StatusCode}";
    }
}
