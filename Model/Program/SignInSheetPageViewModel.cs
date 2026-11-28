namespace Topo.Model.Program
{
    public enum SignInSectionsScope
    {
        Primary,
        WholeGroup,
        ChooseSections
    }

    public class SignInSheetOptions
    {
        public SignInSectionsScope SectionsScope { get; set; } = SignInSectionsScope.Primary;
        public string[] SelectedSections { get; set; } = Array.Empty<string>();
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
        public string UnitId { get; set; } = "";
        public string UnitName { get; set; } = "";
        public string[] SelectedCalendars { get; set; } = new string[0];
        public string CalendarErrorMessage { get; set; } = string.Empty;
        public DateTime CalendarSearchFromDate { get; set; } = DateTime.Now;
        public DateTime CalendarSearchToDate { get; set; } = DateTime.Now;
        public string DateErrorMessage { get; set; } = "";
        public IEnumerable<EventListModel>? Events { get; set; }
        public List<string> AvailableSections { get; set; } = new List<string>();
        public SignInSheetOptions Options { get; set; } = new SignInSheetOptions();
    }
}
