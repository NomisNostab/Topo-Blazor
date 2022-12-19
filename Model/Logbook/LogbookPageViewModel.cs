using Topo.Model.Members;

namespace Topo.Model.Logbook
{
    public class LogbookPageViewModel
    {
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
        public List<MemberListModel> Members { get; set; } = new List<MemberListModel>();
        public bool IncludeLeaders { get; set; } = false;
    }
}
