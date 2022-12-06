using Microsoft.AspNetCore.Components.Forms;
using System.ComponentModel.DataAnnotations;

namespace Topo.Model.Approvals
{
    public class BackupRestorePageViewModel
    {
        [Required(ErrorMessage = "Select unit")]
        public string SelectedUnitId { get; set; } = string.Empty;
        public string SelectedUnitName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public IBrowserFile approvalsFile { get; set; }
    }
}
