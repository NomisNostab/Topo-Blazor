using Topo.Model.Login;
using Newtonsoft.Json;
using System.Text;

namespace Topo.Services
{

    public interface ITerrainAPIService
    {
        public Task<AuthenticationResultModel?> LoginAsync(string? branch, string? username, string? password);
        public Task<GetUserResultModel?> GetUserAsync();
        public Task<GetProfilesResultModel> GetProfilesAsync();
    }

    public class TerrainAPIService : ITerrainAPIService
    {
        private readonly HttpClient _httpClient;
        private readonly StorageService _storageService;
        private readonly string cognitoAddress = "https://cognito-idp.ap-southeast-2.amazonaws.com/";
        private readonly string membersAddress = "https://members.terrain.scouts.com.au/";
        private readonly string eventsAddress = "https://events.terrain.scouts.com.au/";
        private readonly string templatesAddress = "https://templates.terrain.scouts.com.au/";
        private readonly string achievementsAddress = "https://achievements.terrain.scouts.com.au/";
        private readonly string metricsAddress = "https://metrics.terrain.scouts.com.au/";
        private readonly List<string> clientIds = new List<string>
        {
            "6v98tbc09aqfvh52fml3usas3c",
            "5g9rg6ppc5g1pcs5odb7nf7hf9",
            "1u4uajve0lin0ki5n6b61ovva7",
            "21m9o832lp5krto1e8ioo6ldg2"
        };

        public TerrainAPIService(HttpClient httpClient, StorageService storageService)
        {
            _httpClient = httpClient;
            _storageService = storageService;
        }

        public async Task<AuthenticationResultModel?> LoginAsync(string? branch, string? username, string? password)
        {
            AuthenticationResultModel authenticationResultModel = new AuthenticationResultModel();
            var savedClientId = "";
            //TODO Persist savedClientId to browser storage
            var result = "";
            var initiateAuth = new InitiateAuthModel();
            initiateAuth.ClientMetadata = new ClientMetadata();
            initiateAuth.AuthFlow = "USER_PASSWORD_AUTH";
            initiateAuth.AuthParameters = new AuthParameters();
            initiateAuth.AuthParameters.USERNAME = $"{branch}-{username}";
            initiateAuth.AuthParameters.PASSWORD = password;
            if (!string.IsNullOrEmpty(savedClientId))
            {
                initiateAuth.ClientId = savedClientId;
                var content = JsonConvert.SerializeObject(initiateAuth);
                result = await SendRequest(HttpMethod.Post, cognitoAddress, content, "AWSCognitoIdentityProviderService.InitiateAuth");
                var authenticationSuccessResult = JsonConvert.DeserializeObject<AuthenticationSuccessResultModel>(result);
                if (authenticationSuccessResult?.AuthenticationResult != null)
                {
                    authenticationResultModel.AuthenticationSuccessResultModel = authenticationSuccessResult;
                    _storageService.ClientId = savedClientId;
                }
            }
            else
            {
                foreach (var clientId in clientIds)
                {
                    initiateAuth.ClientId = clientId;
                    var content = JsonConvert.SerializeObject(initiateAuth);
                    result = await SendRequest(HttpMethod.Post, cognitoAddress, content, "AWSCognitoIdentityProviderService.InitiateAuth");
                    var authenticationSuccessResult = JsonConvert.DeserializeObject<AuthenticationSuccessResultModel>(result);
                    if (authenticationSuccessResult?.AuthenticationResult != null)
                    {
                        authenticationResultModel.AuthenticationSuccessResultModel = authenticationSuccessResult;
                        _storageService.ClientId = clientId;
                        //TODO Persist savedClientId to browser storage
                        break;
                    }
                }
            }
            if (authenticationResultModel.AuthenticationSuccessResultModel.AuthenticationResult == null)
            {
                var authenticationErrorResultModel = JsonConvert.DeserializeObject<AuthenticationErrorResultModel>(result);
                authenticationResultModel.AuthenticationErrorResultModel = authenticationErrorResultModel;
            }

            return authenticationResultModel;
        }

        public async Task<GetUserResultModel?> GetUserAsync()
        {
            AccessTokenModel accessToken = new AccessTokenModel() { AccessToken = _storageService.AuthenticationResult?.AccessToken };
            var content = JsonConvert.SerializeObject(accessToken);
            var result = await SendRequest(HttpMethod.Post, cognitoAddress, content, "AWSCognitoIdentityProviderService.GetUser");
            var getUserResultModel = JsonConvert.DeserializeObject<GetUserResultModel>(result);

            return getUserResultModel;
        }

        public async Task RefreshTokenAsync()
        {
            if (_storageService.TokenExpiry < DateTime.Now)
            {
                var initiateAuth = new InitiateAuthModel();
                initiateAuth.ClientMetadata = new ClientMetadata();
                initiateAuth.AuthFlow = "REFRESH_TOKEN_AUTH";
                initiateAuth.ClientId = _storageService.ClientId;
                initiateAuth.AuthParameters = new AuthParameters();
                initiateAuth.AuthParameters.REFRESH_TOKEN = _storageService?.AuthenticationResult?.RefreshToken;
                initiateAuth.AuthParameters.DEVICE_KEY = null;
                var content = JsonConvert.SerializeObject(initiateAuth);
                var result = await SendRequest(HttpMethod.Post, cognitoAddress, content, "AWSCognitoIdentityProviderService.InitiateAuth");
                var authenticationResult = JsonConvert.DeserializeObject<AuthenticationSuccessResultModel>(result);
                if (_storageService != null && _storageService.AuthenticationResult != null)
                {
                    _storageService.AuthenticationResult.AccessToken = authenticationResult?.AuthenticationResult?.AccessToken;
                    _storageService.AuthenticationResult.IdToken = authenticationResult?.AuthenticationResult?.IdToken;
                    _storageService.AuthenticationResult.ExpiresIn = authenticationResult?.AuthenticationResult?.ExpiresIn;
                    _storageService.AuthenticationResult.TokenType = authenticationResult?.AuthenticationResult?.TokenType;
                    _storageService.TokenExpiry = DateTime.Now.AddSeconds((authenticationResult?.AuthenticationResult?.ExpiresIn ?? 0) - 60);
                }
            }
        }

        public async Task<GetProfilesResultModel> GetProfilesAsync()
        {
            await RefreshTokenAsync();
            string requestUri = $"{membersAddress}profiles";
            var result = await SendRequest(HttpMethod.Get, requestUri);
            var getProfilesResultModel = DeserializeObject<GetProfilesResultModel>(result);

            return getProfilesResultModel;
        }

        private async Task<string> SendRequest(HttpMethod httpMethod, string requestUri, string content = "", string xAmzTargetHeader = "")
        {
            HttpRequestMessage httpRequest = new HttpRequestMessage(httpMethod, requestUri);
            if (!string.IsNullOrEmpty(content))
                httpRequest.Content = new StringContent(content, Encoding.UTF8, "application/x-amz-json-1.1");
            if (string.IsNullOrEmpty(xAmzTargetHeader))
                httpRequest.Headers.Add("authorization", _storageService?.AuthenticationResult?.IdToken);
            else
                httpRequest.Headers.Add("X-Amz-Target", xAmzTargetHeader);
            httpRequest.Headers.Add("accept", "application/json, text/plain, */*");
            //httpRequest.Headers.Add("X-Amz-User-Agent", "aws-amplify/0.1.x js");

            var response = await _httpClient.SendAsync(httpRequest);
            var responseContent = response.Content.ReadAsStringAsync();
            var result = responseContent.Result;
            if (string.IsNullOrEmpty(xAmzTargetHeader)) // Dont log authorisation requests
            {
                // TODO Logging
                //_logger.LogInformation($"Request: {requestUri}");
                //_logger.LogInformation($"Response: {result}");
            }
            return result;
        }

        private T DeserializeObject<T>(string result)
        {
            try
            {
                return JsonConvert.DeserializeObject<T>(result);
            }
            catch (Exception ex)
            {
                // TODO Logging
                //_logger.LogError($"Error deserialising: {typeof(T)}");
                //_logger.LogError($"String being processed: {result}");
                //_logger.LogError($"Exception message: {ex.Message}");
            }
            return JsonConvert.DeserializeObject<T>("");
        }
    }
}
