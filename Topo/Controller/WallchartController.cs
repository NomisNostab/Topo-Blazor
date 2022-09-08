using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.Wallchart;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class WallchartController : ComponentBase
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
        public IWallchartService _wallchartService { get; set; }

        public WallchartPageViewModel model = new WallchartPageViewModel();

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.Units = _storageService.Units;
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

        internal async Task WallchartReportPdfClick()
        {
            if (string.IsNullOrEmpty(model.UnitId))
                return;

            byte[] report = await WallchartReport(OutputType.PDF);
            var fileName = $"Wallchart_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task WallchartReportXlsxClick()
        {
            if (string.IsNullOrEmpty(model.UnitId))
                return;

            byte[] report = await WallchartReport(OutputType.Excel);
            var fileName = $"Wallchart_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> WallchartReport(OutputType outputType = OutputType.PDF)
        {
            var wallchartItems = await _wallchartService.GetWallchartItems(model.UnitId);

            var groupName = _storageService.GroupName ?? "";
            var unitName = _storageService.UnitName ?? "";
            var section = _storageService.Section;

            var serialisedWallchartItems = JsonConvert.SerializeObject(wallchartItems);

            var report = await _reportService.GetWallchartReport(groupName, section, unitName, outputType, serialisedWallchartItems);
            return report;
        }

    }
}
