namespace Topo.Model.SectionReport
{
    public class SectionReportPageViewModel
    {
        public string GroupName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        // The units ticked: one is a section report, more than one is a group report. Only units the person can see, all in one group.
        public HashSet<string> SelectedUnits { get; set; } = new HashSet<string>();
        public Dictionary<string, string> Calendars { get; set; } = new Dictionary<string, string>();
        public HashSet<string> SelectedCalendars { get; set; } = new HashSet<string>();
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        // Events after now are listed as upcoming up to this date
        // Save each unit's roster when the report is downloaded, so the next report can name who joined or left
        public bool SaveRoster { get; set; } = true;
        public string ErrorMessage { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }
}
