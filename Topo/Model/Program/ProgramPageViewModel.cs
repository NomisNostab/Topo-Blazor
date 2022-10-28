using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;
using Topo.Controller;

namespace Topo.Model.Program
{
    public class ProgramPageViewModel
    {
        public Dictionary<string, string> Calendars { get; set; } = new Dictionary<string, string>();

        [Display(Name = "Unit Calendar")]
        [Required(ErrorMessage = "Select a calendar to show")]
        public string CalendarId { get; set; } = "";
        public IEnumerable<EventListModel>? Events { get; set; }
        public string EventId { get; set; } = "";

        public DateTime CalendarSearchFromDate { get; set; } = DateTime.Now;

        public DateTime CalendarSearchToDate { get; set; } = DateTime.Now;
        public string DateErrorMessage { get; set; } = "";
        public bool IncludeGroupEvents { get; set; } = true;
    }

}
