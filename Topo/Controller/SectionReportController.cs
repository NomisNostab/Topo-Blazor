using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.ReportGeneration;
using Topo.Model.SectionReport;
using Topo.Services;

namespace Topo.Controller
{
    public class SectionReportController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public IProgramService _programService { get; set; }

        [Inject]
        public ISectionReportService _sectionReportService { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public Blazored.LocalStorage.ILocalStorageService _localStorageService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public SectionReportPageViewModel model = new SectionReportPageViewModel();
        public SectionReportResult? Result;
        public bool Busy;
        public string Progress = "";

        // More than one unit ticked makes it a group report
        public bool IsGroup => model.SelectedUnits.Count > 1;
        // What the downloads are called: the report's own title
        public string ReportName => Result?.Report.IsGroup == true ? "Group Report" : "Section Report";
        public bool AllUnitsSelected => model.Units.Count > 0 && model.SelectedUnits.Count == model.Units.Count;

        private static string RosterKey(string unitId) => $"SectionReportRoster_{unitId}";

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
            SetPeriod(ThisQuarterStart(), ThisQuarterStart().AddMonths(3).AddDays(-1));
            if (!string.IsNullOrEmpty(_storageService.UnitId) && model.Units.ContainsKey(_storageService.UnitId))
            {
                ToggleUnit(_storageService.UnitId, true);
            }
        }

        // The calendar named after a unit, if there is one
        private string? CalendarOf(string unitId) => CalendarMatcher.FindForUnit(model.Calendars, model.Units[unitId]);

        internal void ToggleUnit(string unitId, bool selected)
        {
            if (selected)
            {
                model.SelectedUnits.Add(unitId);
            }
            else
            {
                model.SelectedUnits.Remove(unitId);
            }
            Result = null;

            // a unit's calendar is ticked and unticked with it; any other calendar ticked by hand stays as it is
            var calendar = CalendarOf(unitId);
            if (calendar != null)
            {
                if (selected)
                {
                    model.SelectedCalendars.Add(calendar);
                }
                else
                {
                    model.SelectedCalendars.Remove(calendar);
                }
            }
            if (model.SelectedUnits.Count == 1)
            {
                _storageService.UnitId = model.SelectedUnits.First();
            }
        }

        internal void ToggleAllUnits(bool selected)
        {
            foreach (var unitId in model.Units.Keys)
            {
                if (selected != model.SelectedUnits.Contains(unitId))
                {
                    ToggleUnit(unitId, selected);
                }
            }
        }

        internal void ToggleCalendar(string calendarId, bool selected)
        {
            if (selected)
            {
                model.SelectedCalendars.Add(calendarId);
            }
            else
            {
                model.SelectedCalendars.Remove(calendarId);
            }
        }

        public List<List<KeyValuePair<string, string>>> CalendarColumns() => CalendarMatcher.Columns(model.Calendars);

        private static DateTime ThisQuarterStart() => new DateTime(DateTime.Today.Year, (DateTime.Today.Month - 1) / 3 * 3 + 1, 1);

        private void SetPeriod(DateTime from, DateTime to)
        {
            model.From = from;
            model.To = to;
        }

        internal void LastMonth()
        {
            var firstOfThisMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            SetPeriod(firstOfThisMonth.AddMonths(-1), firstOfThisMonth.AddDays(-1));
        }

        internal void ThisQuarter() => SetPeriod(ThisQuarterStart(), ThisQuarterStart().AddMonths(3).AddDays(-1));

        internal void LastQuarter() => SetPeriod(ThisQuarterStart().AddMonths(-3), ThisQuarterStart().AddDays(-1));

        internal void ThisYear() => SetPeriod(new DateTime(DateTime.Today.Year, 1, 1), new DateTime(DateTime.Today.Year, 12, 31));

        internal async Task GenerateClick()
        {
            model.ErrorMessage = "";
            model.Message = "";
            if (model.SelectedUnits.Count == 0)
            {
                model.ErrorMessage = "Please select at least one unit.";
            }
            else if (model.To < model.From)
            {
                model.ErrorMessage = "The to date must be after the from date.";
            }
            if (model.ErrorMessage != "")
            {
                return;
            }

            // in the order the units are listed
            var unitIds = model.Units.Keys.Where(model.SelectedUnits.Contains).ToList();
            var request = new SectionReportRequest
            {
                UnitIds = unitIds,
                IsGroup = IsGroup,
                From = model.From,
                To = model.To
            };
            foreach (var unitId in unitIds)
            {
                // a unit reads its events from its own calendar and from any calendar that is not another ticked unit's
                // (a group calendar, say), so a joint event is listed under each unit that is in the report
                var otherUnitCalendars = unitIds.Where(id => id != unitId).Select(CalendarOf).Where(c => c != null).ToHashSet();
                request.CalendarIdsByUnit[unitId] = model.SelectedCalendars.Where(c => !otherUnitCalendars.Contains(c)).ToArray();
                var saved = await ReadRoster(unitId);
                if (saved != null)
                {
                    request.PreviousRosters[unitId] = saved;
                }
            }

            Busy = true;
            Result = null;
            try
            {
                Result = await _sectionReportService.GetSectionReport(request, step => { Progress = step; InvokeAsync(StateHasChanged); });
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't build the report.";
            }
            finally
            {
                Busy = false;
                Progress = "";
            }
        }

        // A roster that can't be read back (cleared or changed storage) is the same as none having been saved
        private async Task<Topo.Model.SectionReport.RosterSnapshot?> ReadRoster(string unitId)
        {
            try
            {
                return await _localStorageService.GetItemAsync<Topo.Model.SectionReport.RosterSnapshot>(RosterKey(unitId));
            }
            catch (Exception)
            {
                return null;
            }
        }

        internal async Task DownloadClick(OutputType outputType)
        {
            if (Result == null)
            {
                return;
            }

            var report = Result.Report;
            var first = report.Sections.FirstOrDefault();
            var unitName = report.IsGroup ? "Group" : first?.UnitName ?? "";
            // the logo is the section being reported on (the first unit's, for a group)
            var bytes = await _reportService.GetSectionReport(_storageService.GroupName ?? "Group Name", first?.Section ?? "", unitName, outputType, JsonConvert.SerializeObject(report));
            var isPdf = outputType == OutputType.PDF;
            var fileName = $"{(report.IsGroup ? "GroupReport" : "SectionReport")}_{unitName.Replace(' ', '_')}_{report.From:yyyy-MM-dd}_{report.To:yyyy-MM-dd}" + (isPdf ? ".pdf" : ".xlsx");
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, isPdf ? "application/pdf" : "application/vnd.ms-excel", bytes);

            if (model.SaveRoster)
            {
                foreach (var roster in Result.Rosters)
                    await _localStorageService.SetItemAsync(RosterKey(roster.Key), roster.Value);
                model.Message = "Rosters saved: the next report will name who has joined or left since this one.";
            }
        }
    }
}
