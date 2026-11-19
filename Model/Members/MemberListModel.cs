using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace Topo.Model.Members
{
    public class MemberListModel
    {
        public string id { get; set; }
        [Display(Name = "Member Number")]
        public string member_number { get; set; } = string.Empty;
        [Display(Name = "First Name")]
        public string first_name { get; set; } = string.Empty;
        [Display(Name = "Family Name")]
        public string last_name { get; set; } = string.Empty;
        public string status { get; set; } = string.Empty;
        public string age { get; set; } = string.Empty;
        [Display(Name = "Unit Council")]
        public bool unit_council { get; set; }
        [Display(Name = "Patrol")]
        public string patrol_name { get; set; } = string.Empty;
        [Display(Name = "Role")]
        public string patrol_duty { get; set; } = string.Empty;
        public int patrol_order { get; set; } = 0;
        public int isAdultLeader { get; set; } = 0;
        // Unlike every other field here, this is NOT Terrain data - Terrain doesn't expose gender.
        // It's a per-member annotation the leader sets in the Patrol Planner UI and that MembersService
        // persists to this browser's local storage (see SetMemberGenderAsync).
        public string gender { get; set; } = string.Empty;
        public bool selected { get; set; }
        public bool isEligibleJamboree { get; set; } = false;

        // The unit's section, filled in when members are gathered from several units (a sign in sheet
        // for the whole group) so they can be grouped by section. Blank for every other use.
        public string section { get; set; } = string.Empty;

        // A copy, for adding to it without changing the member list that is cached for the unit
        public MemberListModel Clone() => (MemberListModel)MemberwiseClone();
    }
}
