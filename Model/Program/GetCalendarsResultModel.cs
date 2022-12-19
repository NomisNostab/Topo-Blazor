namespace Topo.Model.Program
{
    public class GetCalendarsResultModel
    {
        public string member_id { get; set; } = "";
        public Own_Calendars[] own_calendars { get; set; } = new Own_Calendars[0];
        public Other_Calendars[] other_calendars { get; set; } = new Other_Calendars[0];
    }

    public class Own_Calendars
    {
        public string id { get; set; } = "";
        public string type { get; set; } = "";
        public string title { get; set; } = "";
        public bool selected { get; set; } = false;
        public string section { get; set; } = "";
    }

    public class Other_Calendars
    {
        public string id { get; set; } = "";
        public string type { get; set; } = "";
        public string title { get; set; } = "";
        public bool selected { get; set; } = false;
        public string section { get; set; } = "";
    }
}
