using Blazored.LocalStorage;
using Syncfusion.Blazor.Data;
using Topo.Model.Login;

namespace Topo.Services
{
    public interface ILoginService
    {
        public Task<AuthenticationResultModel?> LoginAsync(string? branch, string? username, string? password);
        public Task GetUserAsync();
        public Task GetProfilesAsync();
        public Dictionary<string, string> GetGroups();
        // Profile/group/leader-status setup shared by a fresh login and a restored session,
        // so the two paths can never drift apart and leave one of them half set up.
        public Task CompleteLoginSetupAsync();
        public Task PersistSessionAsync();
        public Task<bool> TryRestoreSessionAsync();
        public Task ClearPersistedSessionAsync();
    }
    public class LoginService : ILoginService
    {
        private const string SessionStorageKey = "TopoSession";

        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly ILocalStorageService _localStorageService;

        public LoginService(StorageService storageService, ITerrainAPIService terrainAPIService, ILocalStorageService localStorageService)
        {
            _storageService = storageService;
            _terrainAPIService = terrainAPIService;
            _localStorageService = localStorageService;
        }

        public async Task<AuthenticationResultModel?> LoginAsync(string? branch, string? username, string? password)
        {
            var authenticationResultModel = await _terrainAPIService.LoginAsync(branch, username, password);

            _storageService.IsAuthenticated = false;
            if (authenticationResultModel.AuthenticationSuccessResultModel.AuthenticationResult != null)
            {
                _storageService.AuthenticationResult = authenticationResultModel.AuthenticationSuccessResultModel.AuthenticationResult;
                _storageService.IsAuthenticated = true;
                _storageService.TokenExpiry = DateTime.Now.AddSeconds((authenticationResultModel.AuthenticationSuccessResultModel.AuthenticationResult?.ExpiresIn ?? 0) - 60);
            }
            return authenticationResultModel;
        }

        public async Task GetUserAsync()
        {
            var getUserResultModel = await _terrainAPIService.GetUserAsync();
            _storageService.GetUserResult = getUserResultModel;
        }

        public async Task GetProfilesAsync()
        {
            var getProfilesResultModel = await _terrainAPIService.GetProfilesAsync();
            if (_storageService != null)
                _storageService.GetProfilesResult = getProfilesResultModel;
        }

        public Dictionary<string, string> GetGroups()
        {
            var groups = _storageService.GetProfilesResult?.profiles?.Where(p => p.group != null && p.unit != null).Select(p => p.group).ToList();
            // GetProfilesResult can be null here if the last profile fetch failed (e.g. a dead/expired
            // session being restored - see LoginService.TryRestoreSessionAsync), so without this guard
            // the DistinctBy below would throw on a null source instead of just reporting no groups.
            if (groups == null)
                return new Dictionary<string, string>();
            groups = groups.DistinctBy(g => new {g.id}).ToList();
            return groups
                .OrderBy(g => g.name)
                .ToDictionary(g => g.id, g => g.name);
        }

        public async Task CompleteLoginSetupAsync()
        {
            await GetUserAsync();
            await GetProfilesAsync();
            var isLeader = false;
            if (_storageService.GetProfilesResult?.profiles != null && _storageService.GetProfilesResult.profiles.Length > 0)
            {
                _storageService.MemberName = _storageService.GetProfilesResult.profiles[0].member?.name ?? "";
                _storageService.GroupName = _storageService.GetProfilesResult.profiles[0].group?.name ?? "";
                foreach (var profile in _storageService.GetProfilesResult.profiles)
                {
                    if (profile.group != null && profile.group.roles != null && profile.group.roles.Any(x => x == "group-leader" || x.StartsWith("support-leader-group-")))
                    {
                        isLeader = true;
                        break;
                    }
                    if (profile.unit != null && profile.unit.roles != null && profile.unit.roles.Any(x => x == "unit-council"))
                    {
                        isLeader = true;
                        break;
                    }
                }
                _storageService.IsYouthMember = !isLeader;
            }
            _storageService.Groups = GetGroups();
            _storageService.GroupId = _storageService.Groups.Count == 1 ? _storageService.Groups.FirstOrDefault().Key : "";
            _storageService.GroupName = _storageService.Groups.Count == 1 ? _storageService.Groups.FirstOrDefault().Value : "";
        }

        public async Task PersistSessionAsync()
        {
            var session = new PersistedSession
            {
                ClientId = _storageService.ClientId,
                AuthenticationResult = _storageService.AuthenticationResult,
                TokenExpiry = _storageService.TokenExpiry
            };
            await _localStorageService.SetItemAsync(SessionStorageKey, session);
        }

        // Called once, at app startup, before any page has rendered (see App.razor). A missing or
        // dead refresh token just means "not signed in" - it does not surface as an error anywhere.
        public async Task<bool> TryRestoreSessionAsync()
        {
            var session = await _localStorageService.GetItemAsync<PersistedSession>(SessionStorageKey);
            if (session?.AuthenticationResult == null || string.IsNullOrEmpty(session.AuthenticationResult.RefreshToken))
                return false;

            _storageService.ClientId = session.ClientId;
            _storageService.AuthenticationResult = session.AuthenticationResult;
            _storageService.TokenExpiry = session.TokenExpiry;
            _storageService.IsAuthenticated = true;

            try
            {
                // Forces a refresh when the saved access token has expired, using the saved refresh
                // token - the same call every other API method already makes before each request.
                // IMPORTANT: TerrainAPIService's HTTP layer swallows request failures instead of
                // throwing (see SendRequest's catch-all), so a dead/expired refresh token leaves
                // IdToken null rather than raising an exception here - check for that explicitly,
                // the same way CompleteLoginSetupAsync is checked below, instead of relying on an
                // exception that will never actually arrive for this particular failure.
                await _terrainAPIService.RefreshTokenAsync();
                if (string.IsNullOrEmpty(_storageService.AuthenticationResult?.IdToken))
                    throw new InvalidOperationException("Saved session's refresh token was rejected by Terrain.");

                await CompleteLoginSetupAsync();
                if (_storageService.GetProfilesResult?.profiles == null)
                    throw new InvalidOperationException("Could not load a profile for the restored session.");

                return true;
            }
            catch
            {
                // Covers both the explicit checks above and anything else unexpected (e.g. local
                // storage itself failing) - fall back to a clean logged-out state rather than leaving
                // a half-restored, broken session in place.
                await ClearPersistedSessionAsync();
                _storageService.Logout();
                return false;
            }
        }

        public async Task ClearPersistedSessionAsync()
        {
            await _localStorageService.RemoveItemAsync(SessionStorageKey);
        }
    }
}
