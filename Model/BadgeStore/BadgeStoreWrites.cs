using Topo.Model.Program;

namespace Topo.Model.BadgeStore
{
    public enum BadgeStoreWriteOutcome
    {
        // The record was posted and read back
        Written,
        // The store already says this, so nothing was posted
        AlreadyRecorded,
        // There is no record to take back
        NothingToRevoke,
        // The group has no store event yet; creating one needs the leader's go-ahead
        NoStoreEvent,
        // The store event already exists, so another was not created
        StoreAlreadyExists,
        // The store event was created and found again
        StoreCreated,
        Failed
    }

    public class BadgeStoreWriteResult
    {
        public BadgeStoreWriteOutcome Outcome { get; set; }
        // A sentence a screen can show
        public string Message { get; set; } = "";
        public string EventId { get; set; } = "";
        public string CommentId { get; set; } = "";
        public bool Succeeded => Outcome != BadgeStoreWriteOutcome.Failed && Outcome != BadgeStoreWriteOutcome.NoStoreEvent;
    }

    // What a call that changes Terrain came back with. Terrain's rejections can carry no CORS header, so from a browser a
    // failed write can arrive as no readable reply at all (StatusCode 0).
    public class TerrainCallResult
    {
        public bool Success { get; set; }
        public int StatusCode { get; set; }
        public string Body { get; set; } = "";
    }

    // The decisions and request bodies for writing to the store, kept apart from the calls so they can be tested
    public static class BadgeStoreWrites
    {
        private const string TimeFormat = "yyyy-MM-dd'T'HH:mm:ss'.000+00:00'";

        // The body of the call that creates the group's store event. The shape is the one Terrain accepted for a planned
        // event when it was tested against Terrain (see docs/badge-store.md); a planned event with no attendance gives nobody credit.
        public static Dictionary<string, object> CreateEventBody(string groupId, string organiserId, string organiserName)
        {
            var verifier = new Dictionary<string, object> { ["name"] = organiserName, ["contact"] = "", ["type"] = "member" };
            return new Dictionary<string, object>
            {
                ["title"] = BadgeStoreSpec.Title,
                ["description"] = BadgeStoreSpec.Description,
                ["justification"] = "",
                ["organisers"] = new[] { organiserId },
                ["challenge_area"] = "outdoors",
                ["start_datetime"] = BadgeStoreSpec.StartUtc.ToString(TimeFormat),
                ["end_datetime"] = BadgeStoreSpec.EndUtc.ToString(TimeFormat),
                ["event_type"] = new Dictionary<string, object> { ["type"] = "group", ["id"] = groupId },
                ["attendance"] = new Dictionary<string, object>
                {
                    ["leader_member_ids"] = new string[0],
                    ["assistant_member_ids"] = new string[0],
                    ["attendee_member_ids"] = new string[0],
                    ["participant_member_ids"] = new string[0]
                },
                ["schedule_items"] = new object[0],
                ["achievement_pathway_oas_data"] = new Dictionary<string, object>
                {
                    ["award_rule"] = "individual",
                    ["verifier"] = verifier,
                    ["groups"] = new object[0]
                },
                ["achievement_pathway_logbook_data"] = new Dictionary<string, object>
                {
                    ["distance_travelled"] = 0,
                    ["distance_walkabout"] = 0,
                    ["achievement_meta"] = new Dictionary<string, object> { ["stream"] = "", ["branch"] = "" },
                    ["categories"] = new string[0],
                    ["details"] = new Dictionary<string, object> { ["activity_time_length"] = "", ["activity_grade"] = "" },
                    ["title"] = BadgeStoreSpec.Title,
                    ["verifier"] = verifier
                },
                ["review"] = new Dictionary<string, object>
                {
                    ["general_tags"] = new string[0],
                    ["scout_method_elements"] = new Dictionary<string, object>
                    {
                        ["patrol_system"] = true,
                        ["youth_leading_adult_supporting"] = true,
                        ["personal_progression"] = true,
                        ["learn_by_doing"] = true,
                        ["community_involvement"] = true
                    },
                    ["scout_spices_elements"] = new string[0]
                },
                ["uploads"] = new object[0],
                ["equipment_notes"] = "",
                ["additional_notes"] = "",
                ["location"] = "",
                ["iana_timezone"] = "Australia/Sydney",
                ["status"] = "planned"
            };
        }

        // With more than one store event the same one is always written to, so two leaders do not split the records
        public static Event? ChooseWriteEvent(IEnumerable<Event> storeEvents) =>
            storeEvents.OrderBy(e => e.id, StringComparer.Ordinal).FirstOrDefault();

        // True for the event just created, in the leader's organised events (which show it even when the group calendar is off)
        public static bool IsCreatedStoreEvent(Event e, string groupId) =>
            e.group_id == groupId
            && e.title.StartsWith(BadgeStoreSpec.TitlePrefix, StringComparison.OrdinalIgnoreCase)
            && e.start_datetime.ToUniversalTime().Date == BadgeStoreSpec.StartUtc.Date
            && !string.Equals(e.status, "cancelled", StringComparison.OrdinalIgnoreCase);

        // Whether an award would only repeat what the store already says. A request with no presented date does not count
        // as different from a record that has one, because posting it would take the presented date away.
        public static bool IsAlreadyRecorded(BadgeStoreContents contents, string memberId, string awardId, DateTime awarded, DateTime? presented)
        {
            var existing = contents.Records.FirstOrDefault(r => r.MemberId == memberId && r.AwardId == awardId);
            if (existing == null || existing.Awarded.Date != awarded.Date)
                return false;
            return presented == null || existing.Presented?.Date == presented.Value.Date;
        }

        public static bool HasRecord(BadgeStoreContents contents, string memberId, string awardId) =>
            contents.Records.Any(r => r.MemberId == memberId && r.AwardId == awardId);
    }
}
