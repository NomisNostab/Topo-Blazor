namespace Topo.Model.UnitReport
{
    public class UnitReportPageViewModel
    {
        public string GroupName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string UnitId { get; set; } = string.Empty;
        public Dictionary<string, string> Calendars { get; set; } = new Dictionary<string, string>();
        public HashSet<string> SelectedCalendars { get; set; } = new HashSet<string>();
        // Search from the earliest date the report looks at, unless a range is chosen
        public bool AllTime { get; set; } = true;
        public DateTime FromDate { get; set; } = DateTime.Today.AddYears(-1).AddDays(1);
        public DateTime ToDate { get; set; } = DateTime.Today;
        public string ErrorMessage { get; set; } = string.Empty;
        public UnitReportModel? Report { get; set; }
    }
}
