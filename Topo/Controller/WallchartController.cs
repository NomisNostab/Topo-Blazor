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
        public bool Loading;
        public string Progress = "";
        // Laid out again whenever an option changes, from what was read for the unit
        public List<WallchartSheet>? Sheets;
        private List<WallchartItemModel>? _items;

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
            {
                NavigationManager.NavigateTo("index");
                return;
            }

            if (!string.IsNullOrEmpty(_storageService.UnitId))
            {
                model.UnitId = _storageService.UnitId;
                model.UnitName = _storageService.UnitName;
            }

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units;
            model.SuppressLastName = _storageService.SuppressLastName;
            if (!string.IsNullOrEmpty(model.UnitId))
                await Load();
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            _storageService.UnitId = unitId;
            model.UnitId = unitId;
            model.UnitName = _storageService.UnitName;
            await Load();
        }

        // Reads the unit's wallchart once; the preview then follows the options
        private async Task Load()
        {
            Sheets = null;
            _items = null;
            model.ErrorMessage = "";
            if (string.IsNullOrEmpty(model.UnitId))
                return;

            Loading = true;
            try
            {
                _storageService.SuppressLastName = model.SuppressLastName;
                _items = await _wallchartService.GetWallchartItems(model.UnitId, step => { Progress = step; InvokeAsync(StateHasChanged); });
                Recalculate();
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't read the wallchart from Terrain.";
            }
            finally
            {
                Loading = false;
                Progress = "";
            }
        }

        internal void BreakByPatrolChanged() => Recalculate();

        // The items are kept for the unit, so asking again only changes the names to match the setting
        internal async Task SuppressLastNameChanged()
        {
            _storageService.SuppressLastName = model.SuppressLastName;
            if (_items != null)
                _items = await _wallchartService.GetWallchartItems(model.UnitId);
            Recalculate();
        }

        private void Recalculate() => Sheets = _items == null ? null : WallchartLayout.Build(_items, model.BreakByPatrol);

        internal async Task WallchartReportPdfClick() => await Download(OutputType.PDF, "pdf", "application/pdf");

        internal async Task WallchartReportXlsxClick() => await Download(OutputType.Excel, "xlsx", "application/vnd.ms-excel");

        // The report is made from what is on the page, so it is the same as the preview
        private async Task Download(OutputType outputType, string extension, string contentType)
        {
            if (_items == null || _items.Count == 0)
                return;

            try
            {
                var groupName = _storageService.GroupName ?? "";
                var unitName = _storageService.UnitName ?? "";
                var report = await _reportService.GetWallchartReport(groupName, _storageService.Section, unitName, outputType, JsonConvert.SerializeObject(_items), model.BreakByPatrol);
                var fileName = $"Wallchart_{model.UnitName.Replace(' ', '_')}.{extension}";

                // Send the data to JS to actually download the file
                await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, contentType, report);
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't build the report.";
            }
        }
    }
}
