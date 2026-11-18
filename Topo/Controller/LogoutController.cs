using Microsoft.AspNetCore.Components;
using Topo.Model.Index;
using Topo.Services;

namespace Topo.Controller
{
    public class LogoutController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public ILoginService _loginService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }
        protected override async Task OnInitializedAsync()
        {
            // Clears the "Remember Me" entry too - logging out needs to actually log out,
            // not just reset in-memory state that a restore would immediately repopulate.
            await _loginService.ClearPersistedSessionAsync();
            _storageService.Logout();
            NavigationManager.NavigateTo("index");
        }
    }
}
