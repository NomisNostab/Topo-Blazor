using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace Topo.Model.Members
{
    public class MembersPageViewModel
    {
        [Display(Name = "Unit")]
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public List<MemberListModel>? Members { get; set; }
        [Display(Name = "Include Leaders")]
        public bool IncludeLeaders { get; set; } = false;
        public string GroupName { get; set; } = string.Empty;
    }
}
