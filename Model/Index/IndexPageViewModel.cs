using System.ComponentModel.DataAnnotations;

namespace Topo.Model.Index
{
    public class IndexPageViewModel
    {
        public bool IsAuthenticated { get; set; }
        public string Email { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;

        [Required(ErrorMessage = "Select group to access")] 
        public string GroupId { get; set; } = string.Empty;
        public Dictionary<string, string> Groups { get; set; } = new Dictionary<string, string>();
        public string GroupErrorMessage { get; set; } = string.Empty;
    }
}
