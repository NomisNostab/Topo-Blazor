using Topo.Model.AdditionalAwards;
using Topo.Model.Login;
using Topo.Model.Members;
using Topo.Model.OAS;
using Topo.Model.Wallchart;

namespace Topo.Services
{
    public class StorageService
    {
        public event Action OnChange;
        public string ClientId { get; set; } = string.Empty;
        private bool _isAuthenticated;
        private bool _isYouthMember;
        public bool IsAuthenticated
        {
            get { return _isAuthenticated; }
            set
            {
                if (_isAuthenticated != value)
                {
                    _isAuthenticated = value;
                    NotifyStateChanged();
                }
            }
        }
        private void NotifyStateChanged() => OnChange?.Invoke();
        public AuthenticationResult? AuthenticationResult { get; set; }
        public DateTime TokenExpiry { get; set; }
        public GetUserResultModel? GetUserResult { get; set; } = null;
        public GetProfilesResultModel? GetProfilesResult { get; set; }
        public string? MemberName { get; set; }
        public Dictionary<string, string> Groups { get; set; } = new Dictionary<string, string>();
        public string? GroupId { get; set; }
        public string? GroupName { get; set; }
        public string GroupNameDisplay
        {
            get
            {
                return string.IsNullOrEmpty(GroupName) ? "(No Group Selected)" : $"({GroupName})";
            }
        }
        public Dictionary<string, string> Units
        {
            get
            {
                return GetProfilesResult?.profiles?
                    .Where(p => p.group.name == GroupName)
                    .Where(p => p.member.name == MemberName)
                    .Select(p => p.unit)
                    .ToDictionary(p => p?.id?.ToString() ?? "", p => p?.name ?? "");
            }
        }
        public string UnitId { get; set; } = "";
        public string UnitName { get; set; } = "";
        public string Section
        {
            get
            {
                var unit = GetProfilesResult.profiles.FirstOrDefault(u => u.unit.name == UnitName);
                if (unit == null)
                    throw new IndexOutOfRangeException($"No unit found with name {UnitName}. You may not have permissions to this section");
                return unit.unit.section;
            }
        }
        public List<KeyValuePair<string, List<MemberListModel>>> CachedMembers { get; set; } = new List<KeyValuePair<string, List<MemberListModel>>>();
        public Dictionary<string, string> GetCalendarsResult { get; set; } = new Dictionary<string, string>();
        public List<OASTemplate> OASTemplates { get; set; } = new List<OASTemplate>();
        public List<OASStageListModel> OASStages { get; set; } = new List<OASStageListModel>();
        public List<KeyValuePair<string, List<WallchartItemModel>>> CachedWallchartItems { get; set; } = new List<KeyValuePair<string, List<WallchartItemModel>>>();
        public List<AdditionalAwardSpecificationListModel> AdditionalAwardSpecifications { get; set; } = new List<AdditionalAwardSpecificationListModel>();
        public void Logout()
        {
            IsAuthenticated = false;
            GroupName = "";
            CachedMembers = new List<KeyValuePair<string, List<MemberListModel>>>();
        }
        public bool IsYouthMember
        {
            get { return _isYouthMember; }
            set
            {
                if (_isYouthMember != value)
                {
                    _isYouthMember = value;
                    NotifyStateChanged();
                }
            }
        }

    }
}
