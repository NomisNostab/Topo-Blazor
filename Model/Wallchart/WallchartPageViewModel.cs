using System.ComponentModel.DataAnnotations;

namespace Topo.Model.Wallchart
{
    public class WallchartPageViewModel
    {
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();

        [Required(ErrorMessage = "Select unit")]
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
    }
}
