using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace Topo.Model.Milestone
{
    public class MilestonePageViewModel
    {
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();

        [Required(ErrorMessage = "Select unit")] 
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
    }
}
