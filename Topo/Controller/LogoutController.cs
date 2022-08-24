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
        public NavigationManager NavigationManager { get; set; }
        protected override async Task OnInitializedAsync()
        {
            _storageService.IsAuthenticated = false;
            NavigationManager.NavigateTo("index");
        }
    }
}
