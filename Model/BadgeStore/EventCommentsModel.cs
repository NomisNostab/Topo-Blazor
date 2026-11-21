namespace Topo.Model.BadgeStore
{
    // GET events /events/{eventId}/comments, newest first
    public class GetEventCommentsResultModel
    {
        public EventComment[] results { get; set; } = new EventComment[0];
    }

    public class EventComment
    {
        public string id { get; set; } = "";
        public string text { get; set; } = "";
        public DateTime created_datetime { get; set; }
        public string status { get; set; } = "";
        public EventCommentMember member { get; set; } = new EventCommentMember();
        // Not from Terrain: the event the comment was read from, set by whoever reads it, so a record can be undone
        public string EventId { get; set; } = "";
    }

    public class EventCommentMember
    {
        public string id { get; set; } = "";
        public string name { get; set; } = "";
        public string initials { get; set; } = "";
        public string section { get; set; } = "";
    }
}
