using Microsoft.AspNetCore.Components;
using Topo.Model.Login;
using Topo.Services;

namespace Topo.Controller
{
    public class LoginController : ComponentBase
    {
        [Inject]
        public ILoginService _loginService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public LoginPageViewModel model = new LoginPageViewModel();

        internal async Task LogInClick ()
        {
            var authenticationResult = await _loginService.LoginAsync(model.Branch, model.MemberNumber, model.Password);
            model.LoginErrorMessage = "";
            if (authenticationResult != null && authenticationResult.AuthenticationSuccessResultModel.AuthenticationResult != null)
            {
                await _loginService.CompleteLoginSetupAsync();
                if (model.RememberMe)
                {
                    await _loginService.PersistSessionAsync();
                }

                NavigationManager.NavigateTo("index");
            }
            if (authenticationResult != null && authenticationResult.AuthenticationErrorResultModel != null)
            {
                model.LoginErrorMessage = authenticationResult.AuthenticationErrorResultModel.message;
            }
        }
    }
}
