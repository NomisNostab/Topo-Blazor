using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.ReadyToSubmit;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class ReadyToSubmitController : ComponentBase
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
        public IReadyToSubmitService _readyToSubmitService { get; set; }

        public ReadyToSubmitPageViewModel model = new ReadyToSubmitPageViewModel();
        public bool Loading;
        public string Progress = "";
        // Worked out again whenever an option changes, from what was read for the unit
        public ReadyToSubmitReport? Report;
        private ReadyToSubmitData? _data;

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
            {
                NavigationManager.NavigateTo("index");
                return;
            }

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units;
            if (!string.IsNullOrEmpty(_storageService.UnitId))
                await UnitChange(_storageService.UnitId);
        }

        internal async Task UnitChange(ChangeEventArgs e) => await UnitChange(e.Value?.ToString() ?? "");

        internal async Task UnitChange(string unitId)
        {
            model.UnitId = unitId;
            model.ErrorMessage = "";
            Report = null;
            _data = null;
            if (string.IsNullOrEmpty(unitId))
                return;

            _storageService.UnitId = unitId;
            Loading = true;
            try
            {
                var members = await _membersService.GetMembersAsync(unitId);
                var youth = members.Where(m => m.isAdultLeader == 0).ToList();
                _data = await _readyToSubmitService.LoadAsync(youth, _storageService.Section, step => { Progress = step; InvokeAsync(StateHasChanged); });
                Recalculate();
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't read the badges from Terrain.";
            }
            finally
            {
                Loading = false;
                Progress = "";
            }
        }

        // The preview follows the options as they are ticked, without reading Terrain again
        internal void OptionChanged() => Recalculate();

        private void Recalculate()
        {
            if (_data == null)
                return;
            Report = ReadyToSubmitBuilder.Build(new ReadyToSubmitInput
            {
                UnitName = _storageService.UnitName ?? "",
                Section = _storageService.Section,
                Members = _data.Members,
                Templates = _data.Templates,
                LogbookTickOptional = model.LogbookTickOptional,
                MaxTicksShort = Math.Clamp(model.MaxTicksShort, 0, 30),
                UseFullNames = model.UseFullNames,
                Categories = model.SelectedCategories()
            });
        }

        internal async Task PdfClick() => await Download(OutputType.PDF);

        internal async Task XlsxClick() => await Download(OutputType.Excel);

        private async Task Download(OutputType outputType)
        {
            if (Report == null)
                return;
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = Report.UnitName;
            var bytes = await _reportService.GetReadyToSubmitReport(groupName, Report.Section, unitName, outputType, JsonConvert.SerializeObject(Report));
            var isPdf = outputType == OutputType.PDF;
            var fileName = $"ReadyToSubmit_{unitName.Replace(' ', '_')}" + (isPdf ? ".pdf" : ".xlsx");
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, isPdf ? "application/pdf" : "application/vnd.ms-excel", bytes);
        }
    }
}
