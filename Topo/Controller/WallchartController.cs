using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.Wallchart;
using Topo.Model.ReportGeneration;
using Topo.Services;
using System.Reflection;

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

        internal async Task WallchartReportPdfClick()
        {
            if (string.IsNullOrEmpty(model.UnitId))
                return;

            byte[] report = await WallchartReport(OutputType.PDF);
            if (report.Length == 0)
            {
                model.ErrorMessage = "Group life request took too long. Please try Group Life for unit in Terrain first.";
                return;
            }
            var fileName = $"Wallchart_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task WallchartReportXlsxClick()
        {
            if (string.IsNullOrEmpty(model.UnitId))
                return;

            byte[] report = await WallchartReport(OutputType.Excel);
            if (report.Length == 0)
            {
                model.ErrorMessage = "Group life request took too long. Please try Group Life for unit in Terrain first.";
                return;
            }
            var fileName = $"Wallchart_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> WallchartReport(OutputType outputType = OutputType.PDF)
        {
            model.ErrorMessage = "";
            var wallchartItems = await _wallchartService.GetWallchartItems(model.UnitId);
            if (wallchartItems.Count == 0)
            {
                return new byte[0];
            }

            var groupName = _storageService.GroupName ?? "";
            var unitName = _storageService.UnitName ?? "";
            var section = _storageService.Section;

            var serialisedWallchartItems = JsonConvert.SerializeObject(wallchartItems);

            var report = await _reportService.GetWallchartReport(groupName, section, unitName, outputType, serialisedWallchartItems);
            return report;
        }

    }
}
