using System.ComponentModel.DataAnnotations;

namespace Topo.Model.Login
{
    public class LoginPageViewModel
    {
        public Dictionary<string, string> Branches { get; set; } = new Dictionary<string, string>()
        {
            { "ACT", "act" },
            { "NSW", "nsw" },
            { "NT", "nt" },
            { "QLD", "qld" },
            { "SA", "sa" },
            { "TAS", "tas" },
            { "VIC", "vic" },
            { "WA", "wa" }
        };

        [Display(Name = "Branch")]
        public string Branch { get; set; } = string.Empty;

        [Display(Name = "Member Number")]
        public string MemberNumber { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
    }
}
