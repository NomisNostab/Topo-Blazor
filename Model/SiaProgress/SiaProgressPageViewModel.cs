using Topo.Model.Members;

namespace Topo.Model.SiaProgress
{
    public class SiaProgressPageViewModel
    {
        public string GroupName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string UnitId { get; set; } = string.Empty;
        public SiaScope Scope { get; set; } = SiaScope.OneProject;
        // The unit's youth members
        public List<MemberListModel> Members { get; set; } = new List<MemberListModel>();
        public string MemberId { get; set; } = string.Empty;
        public List<SiaProjectChoice> Projects { get; set; } = new List<SiaProjectChoice>();
        public string ProjectId { get; set; } = string.Empty;
        public bool ExcludeCompleted { get; set; }
        public bool IncludeSummary { get; set; }
        public bool DoubleSided { get; set; } = true;
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
