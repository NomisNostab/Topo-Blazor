using Topo.Model.Members;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class MemberNamesTests
    {
        private static MemberListModel Person(string? first, string? last) => new MemberListModel { id = first + last, first_name = first, last_name = last };

        private static List<string> Names(MemberListModel[] people, bool full)
        {
            var names = MemberNames.GetDisplayNames(people, m => m.first_name, m => m.last_name, full);
            return people.Select(m => names[m]).ToList();
        }

        [TestMethod]
        public void AFirstNameThatIsNotSharedIsAllThatIsShown()
        {
            CollectionAssert.AreEqual(new[] { "Ava", "Zoe" }, Names(new[] { Person("Ava", "Lee"), Person("Zoe", "Hall") }, false));
        }

        [TestMethod]
        public void SharedFirstNamesGetJustEnoughOfTheSurnameToTellThemApart()
        {
            var people = new[] { Person("Sam", "Smith"), Person("Sam", "Stone"), Person("Sam", "Brown"), Person("Sam", "Smyth") };
            CollectionAssert.AreEqual(new[] { "Sam Smi", "Sam Sto", "Sam Bro", "Sam Smy" }, Names(people, false));
        }

        [TestMethod]
        public void PeopleWhoStayIdenticalAreNumbered()
        {
            var people = new[] { Person("Zoe", "Hall"), Person("Zoe", "Hall"), Person("Ava", "Lee") };
            CollectionAssert.AreEqual(new[] { "Zoe Hall (1)", "Zoe Hall (2)", "Ava" }, Names(people, false));
            CollectionAssert.AreEqual(new[] { "Zoe Hall (1)", "Zoe Hall (2)", "Ava Lee" }, Names(people, true));
        }

        [TestMethod]
        public void FullNamesAreFirstNameAndSurname()
        {
            CollectionAssert.AreEqual(new[] { "Sam Smith", "Ava Lee" }, Names(new[] { Person("Sam", "Smith"), Person("Ava", "Lee") }, true));
        }

        [TestMethod]
        public void FirstNamesThatDifferOnlyByCaseOrSpacesAreTheSameName()
        {
            var people = new[] { Person("sam ", "Ash"), Person("Sam", "Birch") };
            CollectionAssert.AreEqual(new[] { "sam A", "sam B" }, Names(people, false));
        }

        [TestMethod]
        public void MissingNamesAreTreatedAsBlankNotAnError()
        {
            CollectionAssert.AreEqual(new[] { "Ava", "" }, Names(new[] { Person("Ava", null), Person(null, null) }, false));
        }
    }
}
