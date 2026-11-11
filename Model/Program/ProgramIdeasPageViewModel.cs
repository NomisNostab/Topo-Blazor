namespace Topo.Model.Program
{
    public enum ProgramIdeasRangeOption
    {
        Year,
        SixMonths,
        None,
        UserSet
    }

    public enum ProgramIdeaSortColumn
    {
        ActivityName,
        LastRunDate,
        TimesRun
    }

    public class ProgramIdeasPageViewModel
    {
        public string GroupName { get; set; } = string.Empty;
        public Dictionary<string, string> Calendars { get; set; } = new Dictionary<string, string>();
        public string[] SelectedCalendars { get; set; } = new string[0];
        public string CalendarErrorMessage { get; set; } = string.Empty;

        // With only one calendar there is nothing to choose, so it starts selected
        public void SelectTheOnlyCalendar()
        {
            if (Calendars.Count == 1)
                SelectedCalendars = new[] { Calendars.Keys.First() };
        }

        public ProgramIdeasRangeOption HistoryRange { get; set; } = ProgramIdeasRangeOption.None;
        public DateTime HistoryFromDate { get; set; } = DateTime.Now.AddYears(-1);

        public ProgramIdeasRangeOption HighlightRange { get; set; } = ProgramIdeasRangeOption.SixMonths;
        public DateTime HighlightFromDate { get; set; } = DateTime.Now.AddMonths(-6);

        public List<ProgramIdeaItemModel>? Ideas { get; set; }
        public bool FetchFailed { get; set; }

        public ProgramIdeaSortColumn SortColumn { get; set; } = ProgramIdeaSortColumn.ActivityName;
        public bool SortDescending { get; set; } = false;
    }
}
