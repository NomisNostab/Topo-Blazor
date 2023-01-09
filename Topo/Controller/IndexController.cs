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

        public IndexPageViewModel indexPageViewModel = new IndexPageViewModel();

        protected override void OnInitialized()
        {
            indexPageViewModel.IsAuthenticated = _storageService.IsAuthenticated;
            indexPageViewModel.FullName = _storageService.MemberName ?? "";
            indexPageViewModel.Groups = _storageService.Groups;
            indexPageViewModel.GroupId = _storageService.GroupId ?? "";
        }

        internal void GroupChange(ChangeEventArgs e)
        {
            var groupId = e.Value?.ToString() ?? "";
            _storageService.GroupId = groupId;
            _storageService.GroupName = _storageService.Groups.Where(u => u.Key == groupId).FirstOrDefault().Value;
            indexPageViewModel.GroupId = groupId;
        }

    }
}
