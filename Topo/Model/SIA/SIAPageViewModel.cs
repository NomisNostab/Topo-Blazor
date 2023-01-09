using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;
using Topo.Model.Members;

namespace Topo.Model.SIA
{
    public class SIAPageViewModel
    {
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
        public List<MemberListModel> Members { get; set; } = new List<MemberListModel>();
        public string GroupName { get; set; } = string.Empty;
    }
}
