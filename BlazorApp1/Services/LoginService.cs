using BlazorApp1.Model.Login;

namespace BlazorApp1.Services
{
    public interface ILoginService
    {
        public Task<AuthenticationResultModel?> LoginAsync(string? branch, string? username, string? password);
        public Task GetUserAsync();
        public Task GetProfilesAsync();
        public Dictionary<string, string> GetUnits();
    }
    public class LoginService : ILoginService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;

        public LoginService(StorageService storageService, ITerrainAPIService terrainAPIService)
        {
            _storageService = storageService;
            _terrainAPIService = terrainAPIService;
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

        public Dictionary<string, string> GetUnits()
        {
            return _storageService.GetProfilesResult?.profiles?
                .Where(p => p.member.name == _storageService.MemberName)
                .Select(p => p.unit)
                .ToDictionary(p => p?.id.ToString() ?? "", p => p?.name ?? "");
        }


    }
}
