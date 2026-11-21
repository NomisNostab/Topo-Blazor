using System.Globalization;
using System.Text.Json;

namespace Topo.Model.BadgeStore
{
    // What the store says about one member and one award
    public class BadgeStoreRecord
    {
        public string MemberId { get; set; } = "";
        // The Terrain additional award id, for example camper_award_10 or walkabout_50
        public string AwardId { get; set; } = "";
        public DateTime Awarded { get; set; }
        public DateTime? Presented { get; set; }
        // Who wrote the record, and when
        public string RecordedById { get; set; } = "";
        public DateTime RecordedAt { get; set; }
        public string CommentId { get; set; } = "";
    }

    public class BadgeStoreContents
    {
        public List<BadgeStoreRecord> Records { get; set; } = new List<BadgeStoreRecord>();
        // Comments left out, so a screen can say how many were not used rather than lose them silently
        public int IgnoredUntrustedAuthor { get; set; }
        public int IgnoredUnreadable { get; set; }
        public int IgnoredNewerVersion { get; set; }
    }

    // Turns the comments on the store event into records.
    //
    // Each comment is one JSON object:
    //   {"v":1,"t":"award","m":"<member id>","a":"<award id>","d":"yyyy-MM-dd","p":"yyyy-MM-dd"}   p is optional
    //   {"v":1,"t":"revoke","m":"<member id>","a":"<award id>"}
    // For one member and award the latest comment wins: an award sets the record and a revoke removes it.
    // Only comments written by someone the caller trusts (an adult leader) are used, because at group level everyone in the
    // group can open the event and write comments.
    public static class BadgeStoreReader
    {
        private const string DateFormat = "yyyy-MM-dd";

        public static BadgeStoreContents Resolve(IEnumerable<EventComment> comments, Func<string, bool> isTrustedAuthor)
        {
            var contents = new BadgeStoreContents();
            var latest = new Dictionary<(string member, string award), BadgeStoreRecord?>();

            // Oldest first, with the comment id to break a tie, so the result does not depend on the order Terrain returned them
            foreach (var comment in comments
                .Where(c => !string.Equals(c.status, "deleted", StringComparison.OrdinalIgnoreCase))
                .OrderBy(c => c.created_datetime).ThenBy(c => c.id, StringComparer.Ordinal))
            {
                if (!TryParse(comment.text, out var kind, out var version, out var member, out var award, out var awarded, out var presented))
                {
                    // A well formed record of a newer version is counted apart from a comment that is not a record at all
                    if (version > BadgeStoreSpec.SchemaVersion)
                        contents.IgnoredNewerVersion++;
                    else
                        contents.IgnoredUnreadable++;
                    continue;
                }
                if (!isTrustedAuthor(comment.member.id))
                {
                    contents.IgnoredUntrustedAuthor++;
                    continue;
                }
                var key = (member, award);
                if (kind == "revoke")
                {
                    latest[key] = null;
                    continue;
                }
                latest[key] = new BadgeStoreRecord
                {
                    MemberId = member,
                    AwardId = award,
                    Awarded = awarded,
                    Presented = presented,
                    RecordedById = comment.member.id,
                    RecordedAt = comment.created_datetime,
                    CommentId = comment.id
                };
            }

            contents.Records = latest.Values.Where(r => r != null).Select(r => r!)
                .OrderBy(r => r.MemberId, StringComparer.Ordinal).ThenBy(r => r.AwardId, StringComparer.Ordinal).ToList();
            return contents;
        }

        // The text of the comment for an award or a revoke, in the format above (used when writing, and by the tests)
        public static string AwardText(string memberId, string awardId, DateTime awarded, DateTime? presented)
        {
            var text = $"{{\"v\":{BadgeStoreSpec.SchemaVersion},\"t\":\"award\",\"m\":{Quote(memberId)},\"a\":{Quote(awardId)},\"d\":\"{awarded.ToString(DateFormat, CultureInfo.InvariantCulture)}\"";
            if (presented.HasValue)
                text += $",\"p\":\"{presented.Value.ToString(DateFormat, CultureInfo.InvariantCulture)}\"";
            return text + "}";
        }

        public static string RevokeText(string memberId, string awardId) =>
            $"{{\"v\":{BadgeStoreSpec.SchemaVersion},\"t\":\"revoke\",\"m\":{Quote(memberId)},\"a\":{Quote(awardId)}}}";

        private static string Quote(string value) => JsonSerializer.Serialize(value);

        private static bool TryParse(string text, out string kind, out int version, out string member, out string award, out DateTime awarded, out DateTime? presented)
        {
            kind = member = award = "";
            version = 0;
            awarded = default;
            presented = null;
            try
            {
                using var doc = JsonDocument.Parse(text);
                var root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Object || !root.TryGetProperty("v", out var v) || v.ValueKind != JsonValueKind.Number || !v.TryGetInt32(out version))
                    return false;
                // A newer version may mean something different by the same fields, so it is not read at all
                if (version != BadgeStoreSpec.SchemaVersion)
                    return false;
                kind = Text(root, "t");
                member = Text(root, "m");
                award = Text(root, "a");
                if (member.Length == 0 || award.Length == 0)
                    return false;
                if (kind == "revoke")
                    return true;
                if (kind != "award")
                    return false;
                if (!DateTime.TryParseExact(Text(root, "d"), DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out awarded))
                    return false;
                var p = Text(root, "p");
                if (p.Length > 0)
                {
                    if (!DateTime.TryParseExact(p, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var presentedDate))
                        return false;
                    presented = presentedDate;
                }
                return true;
            }
            catch (JsonException)
            {
                return false;
            }
        }

        private static string Text(JsonElement root, string name) =>
            root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : "";
    }
}
