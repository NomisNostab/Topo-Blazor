using Topo.Model.Login;
using Topo.Model.Members;

namespace Topo.Services
{
    public class StorageService
    {
        public event Action OnChange;
        public string ClientId { get; set; } = string.Empty;
        private bool _isAuthenticated;
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
        public string? GroupName { get; set; }
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
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
    }
}
