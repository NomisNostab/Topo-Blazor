using Topo.Model.Members;

namespace Topo.Model.Approvals
{
    public class ApprovalsPageViewModel
    {
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
        public List<ApprovalsListModel> Approvals { get; set; } = new List<ApprovalsListModel>();
        public DateTime ApprovalSearchFromDate { get; set; } = DateTime.Now;
        public DateTime ApprovalSearchToDate { get; set; } = DateTime.Now;
        public string DateErrorMessage { get; set; } = string.Empty;
        public bool ToBePresented { get; set; }
        public bool IsPresented { get; set; }
        public string GroupName { get; set; } = string.Empty;
    }
}
