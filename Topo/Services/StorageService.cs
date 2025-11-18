using System.Data.SqlTypes;
using System.Reflection;
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
        private string _unitId = string.Empty;

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
                    .Where(p => p.unit != null)
                    .Where(p => p.group.name == (GroupName ?? ""))
                    .Where(p => p.member.name == (MemberName ?? ""))
                    .Select(p => p.unit)
                    .ToDictionary(u => u?.id?.ToString() ?? "", u => u?.name ?? "");
            }
        }
        public string UnitId
        {
            get
            {
                if (string.IsNullOrEmpty(_unitId) && Units.Count == 1)
                {
                    _unitId = Units.First().Key;
                }
                return _unitId;
            }
            set
            {
                _unitId = value;
            }
        }

        public string UnitName
        {
            get
            {
                if (Units != null)
                    return Units.Where(u => u.Key == UnitId).FirstOrDefault().Value;
                else
                    return string.Empty;
            }
        }
        public string Section
        {
            get
            {
                var profile = GetProfilesResult?.profiles?.Where(p => p.group != null && p.unit != null && p.unit.id == _unitId).Select(p => p).FirstOrDefault();
                if (profile == null)
                {
                    throw new IndexOutOfRangeException($"No unit found with id {_unitId}. You may not have permissions to this section");

                }
                else
                {
                    return profile.unit.section;
                }
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
            _unitId = "";
            GetProfilesResult = null;
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

        public bool SuppressLastName { get; set; }

        public string Version = "1.66";
    }
}
