using Topo.Model.Index;
using Topo.Services;
using Microsoft.AspNetCore.Components;
using System.Reflection;

namespace Topo.Controller
{
    public class IndexController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public IndexPageViewModel indexPageViewModel = new IndexPageViewModel();

        protected override void OnInitialized()
        {
            if (!_storageService.IsAuthenticated)
            {
                NavigationManager.NavigateTo("login");
                return;
            }
            indexPageViewModel.IsAuthenticated = _storageService.IsAuthenticated;
            indexPageViewModel.IsYouthMember = _storageService.IsYouthMember;
            indexPageViewModel.FullName = _storageService.MemberName ?? "";
            indexPageViewModel.Groups = _storageService.Groups;
            indexPageViewModel.GroupId = _storageService.GroupId ?? "";
            indexPageViewModel.GroupErrorMessage = string.IsNullOrEmpty(_storageService.GroupId) ? "Please select a group to access" : "";
        }

        internal void GroupChange(ChangeEventArgs e)
        {
            var groupId = e.Value?.ToString() ?? "";
            _storageService.GroupId = groupId;
            _storageService.GroupName = _storageService.Groups.Where(u => u.Key == groupId).FirstOrDefault().Value;
            indexPageViewModel.GroupId = groupId;
            indexPageViewModel.GroupErrorMessage = string.IsNullOrEmpty(_storageService.GroupId) ? "Please select a group to access" : "";
        }

    }
}
