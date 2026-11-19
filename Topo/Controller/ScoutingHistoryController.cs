using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.History;
using Topo.Model.Members;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class ScoutingHistoryController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public IMembersService _membersService { get; set; }

        [Inject]
        public IHistoryService _historyService { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public ScoutingHistoryPageViewModel model = new ScoutingHistoryPageViewModel();

        // Picking another member while one is still loading must not let the older answer overwrite the newer
        private int _loadRequest;

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
            {
                NavigationManager.NavigateTo("index");
                return;
            }

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units ?? new Dictionary<string, string>();
            if (!string.IsNullOrEmpty(_storageService.UnitId))
            {
                model.UnitId = _storageService.UnitId;
                await PopulateMembers();
            }
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            model.UnitId = e.Value?.ToString() ?? "";
            _storageService.UnitId = model.UnitId;
            model.MemberId = "";
            model.History = null;
            model.ErrorMessage = "";
            _loadRequest++;
            await PopulateMembers();
        }

        private async Task PopulateMembers()
        {
            model.Members = string.IsNullOrEmpty(model.UnitId)
                ? new List<MemberListModel>()
                : (await _membersService.GetMembersAsync(model.UnitId))
                    .OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
        }

        internal async Task MemberChange(ChangeEventArgs e)
        {
            model.MemberId = e.Value?.ToString() ?? "";
            model.History = null;
            model.ErrorMessage = "";
            var request = ++_loadRequest;
            if (string.IsNullOrEmpty(model.MemberId))
                return;

            var member = model.Members.FirstOrDefault(m => m.id == model.MemberId);
            ScoutingHistoryModel? history = null;
            var failed = false;
            try
            {
                history = await _historyService.GetScoutingHistory(model.MemberId, $"{member?.first_name} {member?.last_name}".Trim(), member?.patrol_duty ?? "");
            }
            catch (Exception)
            {
                failed = true;
            }

            if (request != _loadRequest)
                return;
            model.History = history;
            if (failed)
                model.ErrorMessage = "Couldn't load that member's history.";
        }

        internal async Task DownloadClick(OutputType outputType)
        {
            if (model.History == null)
                return;

            // the logo is the section the member is in now
            var section = model.History.Sections.LastOrDefault()?.SectionKey ?? "";
            var report = await _reportService.GetScoutingHistoryReport(
                _storageService.GroupName ?? "Group Name", section, _storageService.UnitName ?? "Unit Name", outputType, JsonConvert.SerializeObject(model.History));
            var isPdf = outputType == OutputType.PDF;
            var fileName = $"ScoutingHistory_{model.History.MemberName.Replace(' ', '_')}" + (isPdf ? ".pdf" : ".xlsx");
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, isPdf ? "application/pdf" : "application/vnd.ms-excel", report);
        }
    }
}
