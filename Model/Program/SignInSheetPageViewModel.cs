namespace Topo.Model.Program
{
    public class SignInSheetOptions
    {
        public bool UseFullNames { get; set; } = false;
        public bool UseTimeInOutSignature { get; set; } = true;
        public bool IncludeParticipationColumns { get; set; } = true;
        public bool IncludeRegisteredColumn { get; set; } = false;
        public bool IncludePaidColumn { get; set; } = false;
    }

    public class SignInSheetPageViewModel
    {
        public string GroupName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> Calendars { get; set; } = new Dictionary<string, string>();
        // The units the sheet is for: one for an ordinary event, several for an event the units share. All are in the one group.
        public HashSet<string> SelectedUnits { get; set; } = new HashSet<string>();
        public string[] SelectedCalendars { get; set; } = new string[0];
        public string CalendarErrorMessage { get; set; } = string.Empty;
        public DateTime CalendarSearchFromDate { get; set; } = DateTime.Now;
        public DateTime CalendarSearchToDate { get; set; } = DateTime.Now;
        public string DateErrorMessage { get; set; } = "";
        public IEnumerable<EventListModel>? Events { get; set; }
        public SignInSheetOptions Options { get; set; } = new SignInSheetOptions();
    }
}
