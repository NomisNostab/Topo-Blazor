using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.Login;
using Topo.Model.Members;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{

    public class ProgressController : ComponentBase
    {
        [Inject]
        public IMembersService _membersService { get; set; }

        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public MembersPageViewModel model = new MembersPageViewModel();

        protected async override Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.Units = _storageService.Units;
            model.GroupName = _storageService.GroupNameDisplay;
            if (!string.IsNullOrEmpty(_storageService.UnitId))
            {
                await populateMembers();
            }
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            model.UnitId = unitId;
            _storageService.UnitId = unitId;
            await populateMembers();
        }

        async Task populateMembers()
        {
            if (_storageService.Units != null)
                _storageService.UnitName = _storageService.Units.Where(u => u.Key == _storageService.UnitId).FirstOrDefault().Value;
            var allMembers = await _membersService.GetMembersAsync(_storageService.UnitId);
            model.Members = allMembers.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
            model.UnitName = _storageService.UnitName;
        }

    }
}
