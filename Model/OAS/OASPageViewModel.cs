using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace Topo.Model.OAS
{
    public class OASPageViewModel
    {
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        [Display(Name = "Unit")]
        [Required(ErrorMessage = "Select unit")]
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
        public List<OASStageListModel> Stages { get; set; } = new List<OASStageListModel>();
        [Display(Name = "Stage")]
        public string[] SelectedStages { get; set; } = new string[0];

        [Display(Name = "Hide Completed Members")]
        public bool HideCompletedMembers { get; set; }

        [Display(Name = "Break by Patrol")]
        public bool BreakByPatrol { get; set; } = false;
        public string StagesErrorMessage { get; set; } = "";
        public bool FormatLikeTerrain { get; set; } = false;
    }
}
