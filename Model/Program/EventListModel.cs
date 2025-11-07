using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace Topo.Model.Program
{
    public class EventListModel
    {
        public string Id { get; set; } = string.Empty;
        [Display(Name = "Program Name")]
        public string EventName { get; set; } = "";

        [Display(Name = "Date")]
        public DateTime StartDateTime { get; set; }
        public string StartDateTimeDisplay { get; set; } = string.Empty;
        public DateTime EndDateTime { get; set; }
        public string EndDateTimeDisplay { get; set; } = string.Empty;
        public string DateDisplay => StartDateTime.Date == EndDateTime.Date
            ? $"{StartDateTime:ddd d}"
            : $"{StartDateTime:ddd d} - {EndDateTime:ddd d}";
        public string StartFinishDisplay { get; set; } = string.Empty;

        [Display(Name = "Challenge Area")]
        public string ChallengeArea { get; set; } = string.Empty;
        public string EventDisplay { get; set; } = string.Empty;

        [Display(Name = "Date Time")]
        public string EventDate => StartDateTime.Date == EndDateTime.Date 
            ? $"{StartDateTime:dd/MM/yy HH:mm} - {EndDateTime:HH:mm}"
            : $"{StartDateTime:dd/MM/yy HH:mm} - {EndDateTime:HH:mm} +{EndDateTime.DayOfYear - StartDateTime.DayOfYear}";

        public List<EventAttendance> attendees = new List<EventAttendance>();

        [Display(Name = "Status")]
        public string EventStatus { get; set; } = string.Empty;
        public string EventScope { get; set; } = string.Empty;
        public string Organiser { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public string Lead { get; set; } = string.Empty;
        public string Assist { get; set; } = string.Empty;
    }

    public class EventAttendance
    {
        public string id { get; set; } = "";
        public string first_name { get; set; } = "";
        public string last_name { get; set; } = "";
        public string member_number { get; set; } = "";
        public string patrol_name { get; set; } = "";
        public bool isAdultMember { get; set; }
        public bool attended { get; set; }
        public string pal { get; set; } = string.Empty;
    }
}
