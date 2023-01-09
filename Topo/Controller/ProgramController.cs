using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using Topo.Model.Program;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class ProgramController : ComponentBase
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
        public IMembersService _membersService { get; set; }

        public ProgramPageViewModel model = new ProgramPageViewModel();
        private string groupCalendarId = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.GroupName = _storageService.GroupNameDisplay;
            await _programService.GetCalendars();
            model.Calendars = _storageService.Units;
            model.CalendarSearchFromDate = DateTime.Now;
            model.CalendarSearchToDate = DateTime.Now.AddMonths(4);
            model.DateErrorMessage = "";
            groupCalendarId = _storageService.GroupId ?? "";
        }

        internal void CalendarChange(ChangeEventArgs e)
        {
            var calendarId = e.Value?.ToString() ?? "";
            model.CalendarId = calendarId;
        }

        internal async Task ShowUnitCalendarClick()
        {
            model.DateErrorMessage = "";
            if (model.CalendarSearchToDate < model.CalendarSearchFromDate)
            {
                model.DateErrorMessage = "The search to date must be after the search from date.";
                return;
            }
            if (!string.IsNullOrEmpty(model.CalendarId))
            {
                _storageService.UnitName = _storageService.Units.Where(u => u.Key == model.CalendarId)?.FirstOrDefault().Value ?? "";
                _storageService.UnitId = model.CalendarId;
                await _programService.SetCalendar(model.CalendarId);
                var events = await _programService.GetEventsForDates(model.CalendarSearchFromDate, model.CalendarSearchToDate);
                await _programService.ResetCalendar();
                if (model.IncludeGroupEvents)
                {
                    await _programService.SetCalendar(groupCalendarId);
                    var groupEvents = await _programService.GetEventsForDates(model.CalendarSearchFromDate, model.CalendarSearchToDate);
                    await _programService.ResetCalendar();
                    events = events.Concat(groupEvents).OrderBy(e => e.StartDateTime).ToList();
                }
                model.Events = events;
            }
        }

        public async Task SignInSheetClick(string eventId)
        {
            var eventName = model.Events?.Where(e => e.Id == eventId).FirstOrDefault()?.EventDisplay ?? "";
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = _storageService.UnitName ?? "Unit Name";
            var section = _storageService.Section;

            var members = await _membersService.GetMembersAsync(model.CalendarId);
            var serialisedSortedMemberList = JsonConvert.SerializeObject(members.OrderBy(m => m.first_name).ThenBy(m => m.last_name));
            var report = await _reportService.GetSignInSheetReport(groupName, section, unitName, OutputType.PDF, serialisedSortedMemberList, eventName);

            var fileName = $"SignInSheet_{unitName.Replace(' ', '_')}_{eventName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task EventAttendanceClick(string eventId)
        {
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = _storageService.UnitName ?? "Unit Name";
            var section = _storageService.Section;

            var eventListModel = await _programService.GetAttendanceForEvent(eventId);
            var serialisedEventListModel = JsonConvert.SerializeObject(eventListModel);
            var report = await _reportService.GetEventAttendanceReport(groupName, section, unitName, OutputType.Excel, serialisedEventListModel);

            var fileName = $"Attendance_{unitName.Replace(' ', '_')}_{eventListModel.EventDisplay.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        internal async Task AttendanceReportPdfClick()
        {
            byte[] report = await AttendanceReport(OutputType.PDF);
            var fileName = $"Attendance_Report_{_storageService.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task AttendanceReportXlsxClick()
        {
            byte[] report = await AttendanceReport(OutputType.Excel);
            var fileName = $"Attendance_Report_{_storageService.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> AttendanceReport(OutputType outputType = OutputType.PDF)
        {
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = _storageService.UnitName ?? "Unit Name";
            var section = _storageService.Section;

            var attendanceReportData = await _programService.GenerateAttendanceReportData(model.CalendarSearchFromDate, model.CalendarSearchToDate, model.CalendarId, model.IncludeGroupEvents ? groupCalendarId : "");
            var serialisedAttendanceReportData = JsonConvert.SerializeObject(attendanceReportData);
            var report = await _reportService.GetAttendanceReport(groupName, section, unitName, outputType, serialisedAttendanceReportData, model.CalendarSearchFromDate, model.CalendarSearchToDate);

            return report;
        }
    }

}
