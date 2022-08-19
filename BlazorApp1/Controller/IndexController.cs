using BlazorApp1.Model.Index;
using BlazorApp1.Services;
using Microsoft.AspNetCore.Components;

namespace BlazorApp1.Controller
{
    public class IndexController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        public IndexPageViewModel indexPageViewModel = new IndexPageViewModel();

        protected override async Task OnInitializedAsync()
        {
            indexPageViewModel.IsAuthenticated = _storageService.IsAuthenticated;
            indexPageViewModel.FullName = _storageService.MemberName ?? "";
        }
    }
}
