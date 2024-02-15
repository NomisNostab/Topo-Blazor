using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.Milestone;
using Topo.Model.OAS;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class MilestoneController : ComponentBase
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
        public IMilestoneService _milestoneService { get; set; }

        public MilestonePageViewModel model = new MilestonePageViewModel();

        protected override void OnInitialized()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units;
            model.UnitId = _storageService.UnitId;
            model.UnitName = _storageService.UnitName;
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            model.UnitId = unitId;
            _storageService.UnitId = model.UnitId;
            if (_storageService.Units != null)
                _storageService.UnitName = _storageService.Units.Where(u => u.Key == model.UnitId).FirstOrDefault().Value;
            model.UnitName = _storageService.UnitName;
        }

        internal async Task MilestoneReportPdfClick()
        {
            if (string.IsNullOrEmpty(model.UnitId))
                return;

            byte[] report = await MilestoneReport(OutputType.PDF);
            var fileName = $"Milestone_Report_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task MilestoneReportXlsxClick()
        {
            if (string.IsNullOrEmpty(model.UnitId))
                return;

            byte[] report = await MilestoneReport(OutputType.Excel);
            var fileName = $"Milestone_Report_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> MilestoneReport(OutputType outputType = OutputType.PDF)
        {
            var milestoneSummaries = await _milestoneService.GetMilestoneSummaries(model.UnitId);

            var groupName = _storageService.GroupName ?? "";
            var unitName = _storageService.UnitName ?? "";
            var section = _storageService.Section;

            var serialisedMilestoneSummaries = JsonConvert.SerializeObject(milestoneSummaries);

            var report = await _reportService.GetMilestoneReport(groupName, section, unitName, outputType, serialisedMilestoneSummaries);
            return report;
        }
    }
}
