using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.ReportGeneration;
using Topo.Model.UnitReport;
using Topo.Services;

namespace Topo.Controller
{
    public class UnitReportController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public IProgramService _programService { get; set; }

        [Inject]
        public IUnitReportService _unitReportService { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public UnitReportPageViewModel model = new UnitReportPageViewModel();
        // Where "All time" starts searching from
        private static readonly DateTime EarliestDate = new DateTime(2018, 1, 1);

        public bool Busy;
        public string Progress = "";

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
            {
                NavigationManager.NavigateTo("index");
                return;
            }

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units ?? new Dictionary<string, string>();
            model.Calendars = await _programService.GetCalendars();
            if (!string.IsNullOrEmpty(_storageService.UnitId))
            {
                SelectUnit(_storageService.UnitId);
            }
        }

        internal void UnitChange(ChangeEventArgs e)
        {
            SelectUnit(e.Value?.ToString() ?? "");
            model.Report = null;
        }

        private void SelectUnit(string unitId)
        {
            model.UnitId = unitId;
            _storageService.UnitId = unitId;

            // use the calendar named after the unit when there is one
            var match = CalendarMatcher.FindForUnit(model.Calendars, _storageService.UnitName);
            if (match != null)
                model.SelectedCalendars = new HashSet<string> { match };
        }

        internal void AllTimeChange(ChangeEventArgs e)
        {
            model.AllTime = (bool)(e.Value ?? false);
        }

        public List<List<KeyValuePair<string, string>>> CalendarColumns() => CalendarMatcher.Columns(model.Calendars);

        internal void ToggleCalendar(string calendarId, bool selected)
        {
            if (selected)
                model.SelectedCalendars.Add(calendarId);
            else
                model.SelectedCalendars.Remove(calendarId);
        }

        internal async Task GenerateClick()
        {
            model.ErrorMessage = "";
            if (string.IsNullOrEmpty(model.UnitId))
                model.ErrorMessage = "Please select a unit.";
            else if (model.SelectedCalendars.Count == 0)
                model.ErrorMessage = "Please select at least one calendar.";
            else if (!model.AllTime && model.ToDate < model.FromDate)
                model.ErrorMessage = "The to date must be after the from date.";
            if (model.ErrorMessage != "")
                return;

            Busy = true;
            model.Report = null;
            try
            {
                var from = model.AllTime ? EarliestDate : model.FromDate;
                var to = model.AllTime ? DateTime.Today : model.ToDate;
                var report = await _unitReportService.GetUnitReport(model.UnitId, from, to, model.SelectedCalendars.ToArray(),
                    step => { Progress = step; InvokeAsync(StateHasChanged); });
                report.AllTime = model.AllTime;
                model.Report = report;
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't build the unit report.";
            }
            finally
            {
                Busy = false;
                Progress = "";
            }
        }

        internal async Task DownloadClick(OutputType outputType)
        {
            if (model.Report == null)
                return;

            var section = model.Report.Section;
            var report = await _reportService.GetUnitReport(_storageService.GroupName ?? "Group Name", section, model.Report.UnitName, outputType, JsonConvert.SerializeObject(model.Report));
            var isPdf = outputType == OutputType.PDF;
            var period = model.Report.AllTime ? "AllTime" : $"{model.Report.From:yyyy-MM-dd}_{model.Report.To:yyyy-MM-dd}";
            var fileName = $"UnitHistory_{model.Report.UnitName.Replace(' ', '_')}_{period}" + (isPdf ? ".pdf" : ".xlsx");
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, isPdf ? "application/pdf" : "application/vnd.ms-excel", report);
        }
    }
}
