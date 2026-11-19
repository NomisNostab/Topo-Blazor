using Topo.Model.Program;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class ProgramIdeasPage
    {
        [TestMethod]
        public void TheOnlyCalendarStartsSelected()
        {
            var model = new ProgramIdeasPageViewModel { Calendars = new Dictionary<string, string> { ["cal-1"] = "Scout Unit" } };

            model.SelectTheOnlyCalendar();

            CollectionAssert.AreEqual(new[] { "cal-1" }, model.SelectedCalendars);
        }

        [TestMethod]
        public void SeveralCalendarsOrNoneAreLeftForThePersonToChoose()
        {
            var several = new ProgramIdeasPageViewModel { Calendars = new Dictionary<string, string> { ["a"] = "A", ["b"] = "B" } };
            several.SelectTheOnlyCalendar();
            Assert.AreEqual(0, several.SelectedCalendars.Length);

            var none = new ProgramIdeasPageViewModel();
            none.SelectTheOnlyCalendar();
            Assert.AreEqual(0, none.SelectedCalendars.Length);
        }
    }
}
