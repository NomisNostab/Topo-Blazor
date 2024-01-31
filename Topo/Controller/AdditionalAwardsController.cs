using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.AdditionalAwards;
using Topo.Model.ReportGeneration;
using Topo.Model.SIA;
using Topo.Services;

namespace Topo.Controller
{
    public class AdditionalAwardsController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        [Inject]
        public IMembersService _membersService { get; set; }

        [Inject]
        public IAdditionalAwardService _additionalAwardsService { get; set; }

        public AdditionalAwardsPageViewModel model = new AdditionalAwardsPageViewModel();

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units;
            if (!string.IsNullOrEmpty(_storageService.UnitId))
            {
                await UnitChange(_storageService.UnitId);
            }
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            await UnitChange(unitId);
        }

        internal async Task UnitChange(string unitId)
        {
            model.UnitId = unitId;
            _storageService.UnitId = model.UnitId;
            if (_storageService.Units != null)
                _storageService.UnitName = _storageService.Units.Where(u => u.Key == model.UnitId).FirstOrDefault().Value;
            var allMembers = await _membersService.GetMembersAsync(model.UnitId);
            model.Members = allMembers.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
            model.UnitName = _storageService.UnitName;
        }

        internal async Task SelectAllChange(ChangeEventArgs e)
        {
            var selectAll = (bool)e.Value;
            foreach (var member in model.Members)
            {
                member.selected = selectAll;
            }
        }

        internal async Task AwardsReportPdfClick()
        {
            byte[] report = await AwardsReport(OutputType.PDF);
            var fileName = $"Additional_Awards_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task AwardsReportXlsxClick()
        {
            byte[] report = await AwardsReport(OutputType.Excel);
            var fileName = $"Additional_Awards_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        internal async Task<byte[]> AwardsReport(OutputType outputType = OutputType.PDF)
        {
            if (!model.Members.Any(m => m.selected))
            {
                foreach (var member in model.Members)
                {
                    member.selected = true;
                }
            }
            var selectedMembers = model.Members.Where(m => m.selected).Select(m => m.id).ToList();

            var memberKVP = new List<KeyValuePair<string, string>>();
            foreach (var member in selectedMembers)
            {
                var memberName = model.Members.Where(m => m.id == member).Select(m => m.first_name + " " + m.last_name).FirstOrDefault();
                memberKVP.Add(new KeyValuePair<string, string>(member, memberName ?? ""));
            }

            var groupName = _storageService.GroupName ?? "";
            var unitName = _storageService.UnitName ?? "";
            var section = _storageService.Section;

            var reportData = await _additionalAwardsService.GenerateAdditionalAwardReportData(model.UnitId, memberKVP);
            var serialisedReportData = JsonConvert.SerializeObject(reportData);

            var report = await _reportService.GetAdditionalAwardsReport(groupName, section, unitName, outputType, serialisedReportData);
            return report;
        }

    }
}
