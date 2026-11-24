using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.Members;
using Topo.Model.Program;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class SignInSheetController : ComponentBase, IDisposable
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public IProgramService _programService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        [Inject]
        public Blazored.LocalStorage.ILocalStorageService _localStorageService { get; set; }

        public ElementReference _select2Reference;

        public SignInSheetPageViewModel model = new SignInSheetPageViewModel();

        // Live preview of the sheet for the chosen event. Names, columns and layout are applied on the
        // page; only the units ticked change who is invited, so only they refetch.
        public string EventId = "";
        public List<MemberListModel>? PreviewMembers;
        public bool PreviewFailed;
        private int _previewRequest;

        private const string OptionsKey = "SignInSheetOptions";

        public bool AllUnitsSelected => model.Units.Count > 0 && model.SelectedUnits.Count == model.Units.Count;

        private DotNetObjectReference<SignInSheetController>? _selfReference;
        // set when the calendar selection was changed in code, so the select2 box can be refreshed to show it
        private bool _syncCalendarSelection;
        private bool _loadingEvents;
        private bool _reloadEvents;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                await JS.InvokeVoidAsync("BindSelect2");
                _selfReference = DotNetObjectReference.Create(this);
                await JS.InvokeVoidAsync("bindSelect2Change", _select2Reference, _selfReference);
            }

            if (_syncCalendarSelection)
            {
                _syncCalendarSelection = false;
                await JS.InvokeVoidAsync("setSelectedValues", _select2Reference, model.SelectedCalendars);
            }
        }

        public void Dispose()
        {
            _selfReference?.Dispose();
        }

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
            // upcoming events: from today to the end of the four months starting this quarter
            var quarter = (DateTime.Now.Month + 2) / 3;
            var quarterStartMonth = (quarter - 1) * 3 + 1;
            model.CalendarSearchFromDate = DateTime.Today;
            model.CalendarSearchToDate = new DateTime(DateTime.Now.Year, quarterStartMonth, 1).AddMonths(4).AddDays(-1);
            if (!string.IsNullOrEmpty(_storageService.UnitId) && model.Units.ContainsKey(_storageService.UnitId))
            {
                SetUnit(_storageService.UnitId, true);
            }

            var savedOptions = await _localStorageService.GetItemAsync<SignInSheetOptions>(OptionsKey);
            if (savedOptions != null)
            {
                model.Options = savedOptions;
            }

            await LoadEventsAsync();
            await RefreshPreview();
        }

        // Saved per browser so the options don't have to be re-picked every time.
        private async Task SaveOptions()
        {
            await _localStorageService.SetItemAsync(OptionsKey, model.Options);
        }

        internal async Task ToggleUnit(string unitId, bool selected)
        {
            SetUnit(unitId, selected);
            await UnitsChanged();
        }

        internal async Task ToggleAllUnits(bool selected)
        {
            foreach (var unitId in model.Units.Keys)
                SetUnit(unitId, selected);
            await UnitsChanged();
        }

        // Ticks or unticks a unit, and its calendar with it
        private void SetUnit(string unitId, bool selected)
        {
            if (selected)
            {
                model.SelectedUnits.Add(unitId);
            }
            else
            {
                model.SelectedUnits.Remove(unitId);
            }
            SetCalendarOfUnit(unitId, selected);
        }

        private async Task UnitsChanged()
        {
            // one unit on its own is the unit the rest of Topo works on
            if (model.SelectedUnits.Count == 1)
            {
                _storageService.UnitId = model.SelectedUnits.First();
            }
            await LoadEventsAsync();
            await RefreshPreview();
        }

        // The units ticked, in the order they are listed
        private List<string> SelectedUnitIds() => model.Units.Keys.Where(model.SelectedUnits.Contains).ToList();

        // A unit's calendar (the one named after it) is ticked and unticked with it; calendars picked by hand stay as they are
        private void SetCalendarOfUnit(string unitId, bool selected)
        {
            var calendarId = FindCalendarForUnit(model.Units[unitId]);
            if (calendarId == null)
            {
                return;
            }

            var calendars = model.SelectedCalendars.Where(c => c != calendarId).ToList();
            if (selected)
            {
                calendars.Add(calendarId);
            }
            model.SelectedCalendars = calendars.ToArray();
            _syncCalendarSelection = true;
        }

        // A calendar named the same as the unit; failing that, one whose name contains the unit's name
        // (or the other way round).
        private string? FindCalendarForUnit(string? unitName)
        {
            if (string.IsNullOrWhiteSpace(unitName))
                return null;

            var name = unitName.Trim();
            var calendars = model.Calendars.Select(c => (id: c.Key, title: c.Value.Trim())).ToList();
            return (calendars.FirstOrDefault(c => string.Equals(c.title, name, StringComparison.OrdinalIgnoreCase)).id
                ?? calendars.FirstOrDefault(c => c.title.Contains(name, StringComparison.OrdinalIgnoreCase)).id
                ?? calendars.FirstOrDefault(c => c.title.Length > 0 && name.Contains(c.title, StringComparison.OrdinalIgnoreCase)).id);
        }

        // Called from the calendar picker when the person changes the selection.
        [JSInvokable]
        public async Task OnCalendarsChanged()
        {
            model.SelectedCalendars = (await JS.InvokeAsync<List<string>>("getSelectedValues", _select2Reference)).ToArray();
            await LoadEventsAsync();
            StateHasChanged();
        }

        // Loads the events for the chosen calendars and dates. Reloads happen on every change, but only one
        // load runs at a time (each one has to switch the calendars over in Terrain and back); a change made
        // while one is running triggers one more load afterwards.
        internal async Task LoadEventsAsync()
        {
            if (_loadingEvents)
            {
                _reloadEvents = true;
                return;
            }

            _loadingEvents = true;
            try
            {
                do
                {
                    _reloadEvents = false;
                    await LoadEventsOnce();
                } while (_reloadEvents);
            }
            finally
            {
                _loadingEvents = false;
            }
        }

        private async Task LoadEventsOnce()
        {
            model.CalendarErrorMessage = "";
            model.DateErrorMessage = "";
            if (model.SelectedCalendars.Length == 0)
            {
                model.CalendarErrorMessage = "Please select at least one calendar";
                model.Events = null;
            }
            else if (model.CalendarSearchToDate < model.CalendarSearchFromDate)
            {
                model.DateErrorMessage = "The search to date must be after the search from date.";
                model.Events = null;
            }
            else
            {
                await _programService.SetCalendar(model.SelectedCalendars);
                var events = await _programService.GetEventsForDates(model.CalendarSearchFromDate, model.CalendarSearchToDate);
                await _programService.ResetCalendar();
                model.Events = events.OrderBy(e => e.StartDateTime).ToList();
            }

            // the previously chosen event may no longer be in the list
            if (EventId != "" && (model.Events == null || !model.Events.Any(e => e.Id == EventId)))
            {
                EventId = "";
                await RefreshPreview();
            }
        }

        internal async Task EventChange(ChangeEventArgs e)
        {
            EventId = e.Value?.ToString() ?? "";
            await RefreshPreview();
        }

        private async Task RefreshPreview()
        {
            var request = ++_previewRequest;
            PreviewFailed = false;
            // the sheet is for a unit; with no event it has a blank title
            if (model.SelectedUnits.Count == 0)
            {
                PreviewMembers = null;
                return;
            }

            List<MemberListModel>? members = null;
            try
            {
                var unitIds = SelectedUnitIds();
                members = string.IsNullOrEmpty(EventId)
                    ? await _programService.GetMembersForUnits(unitIds)
                    : await _programService.GetInviteesForEvent(EventId, unitIds);
            }
            catch (Exception)
            {
                // reported below as a message
            }

            // ignore the answer if the user has picked something else in the meantime
            if (request != _previewRequest)
                return;
            PreviewMembers = members;
            PreviewFailed = members == null;
        }

        internal async Task FullNamesChange(ChangeEventArgs e)
        {
            model.Options.UseFullNames = (bool)(e.Value ?? false);
            await SaveOptions();
        }

        internal async Task AttendanceModeChange(ChangeEventArgs e)
        {
            model.Options.UseTimeInOutSignature = e.Value?.ToString() == "timeinout";
            await SaveOptions();
        }

        internal async Task RegisteredColumnChange(ChangeEventArgs e)
        {
            model.Options.IncludeRegisteredColumn = (bool)(e.Value ?? false);
            await SaveOptions();
        }

        internal async Task PaidColumnChange(ChangeEventArgs e)
        {
            model.Options.IncludePaidColumn = (bool)(e.Value ?? false);
            await SaveOptions();
        }

        internal async Task ParticipationColumnsChange(ChangeEventArgs e)
        {
            model.Options.IncludeParticipationColumns = (bool)(e.Value ?? false);
            await SaveOptions();
        }

        internal async Task DownloadClick(OutputType outputType)
        {
            if (PreviewMembers == null)
                return;

            // no event chosen: the sheet has no event name
            var eventName = string.IsNullOrEmpty(EventId)
                ? ""
                : model.Events?.Where(e => e.Id == EventId).FirstOrDefault()?.EventDisplay ?? "";
            var groupName = _storageService.GroupName ?? "Group Name";
            // the units' names (the sheet is for all of them) and, when they are all one section, its logo
            var units = SelectedUnitIds();
            var unitName = string.Join(", ", units.Select(id => model.Units[id]));
            var sections = units.Select(id => (_storageService.UnitSections ?? new Dictionary<string, string>()).GetValueOrDefault(id, "")).Distinct().ToList();
            var section = sections.Count == 1 ? sections[0] : "";

            var serialisedSortedMemberList = JsonConvert.SerializeObject(PreviewMembers.OrderBy(m => m.first_name).ThenBy(m => m.last_name));
            var report = await _reportService.GetSignInSheetReport(groupName, section, unitName, outputType, serialisedSortedMemberList, eventName, model.Options);

            var isPdf = outputType == OutputType.PDF;
            // the units' names and the event's date hold commas and slashes, which don't belong in a file name
            var nameInFile = System.Text.RegularExpressions.Regex.Replace(unitName, @"[\s,]+", "_");
            var eventInFile = System.Text.RegularExpressions.Regex.Replace(eventName, @"[\s,/\\:*?""<>|]+", "_");
            var fileName = (string.IsNullOrEmpty(eventName) ? $"SignInSheet_{nameInFile}" : $"SignInSheet_{nameInFile}_{eventInFile}") + (isPdf ? ".pdf" : ".xlsx");
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, isPdf ? "application/pdf" : "application/vnd.ms-excel", report);
        }
    }
}
