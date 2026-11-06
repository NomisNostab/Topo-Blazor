namespace Topo.Model.SectionReport
{
    public class SectionReportPageViewModel
    {
        // The scope chosen in the unit box: one unit's id, or this for every unit the person can see
        public const string WholeGroup = "group";

        public string GroupName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string UnitId { get; set; } = string.Empty;
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
