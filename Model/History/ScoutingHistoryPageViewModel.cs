using Topo.Model.Members;

namespace Topo.Model.History
{
    public class ScoutingHistoryPageViewModel
    {
        public string GroupName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string UnitId { get; set; } = string.Empty;
        public List<MemberListModel> Members { get; set; } = new List<MemberListModel>();
        public string MemberId { get; set; } = string.Empty;
        public ScoutingHistoryModel? History { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
