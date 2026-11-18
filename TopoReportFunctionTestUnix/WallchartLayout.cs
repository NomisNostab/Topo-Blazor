using Topo.Model.Wallchart;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class WallchartLayoutTests
    {
        private static WallchartItemModel Scout(string name, string patrol = "") => new WallchartItemModel { MemberName = name, MemberPatrol = patrol };

        private static WallchartRow RowOf(WallchartItemModel scout) => WallchartLayout.Build(new[] { scout }, false).Single().Rows.Single();

        // The cells of a row by the column they sit in, so a spanning cell is found at the first of its columns
        private static WallchartCell At(WallchartRow row, string column)
        {
            var wanted = Array.IndexOf(WallchartLayout.Columns, column);
            Assert.IsTrue(wanted >= 0, column);
            var position = 0;
            foreach (var cell in row.Cells)
            {
                if (position == wanted)
                    return cell;
                position += cell.Span;
            }
            throw new AssertFailedException($"{column} is inside a spanning cell");
        }

        [TestMethod]
        public void EveryRowCoversExactlyTheColumnsTheHeadingsDo()
        {
            Assert.AreEqual(40, WallchartLayout.Columns.Length);
            Assert.AreEqual(40, WallchartLayout.Groups.Sum(g => g.Span));
            Assert.AreEqual(9, WallchartLayout.OutdoorGroups.Sum(g => g.Span));

            var bare = RowOf(Scout("Amy"));
            var awarded = Scout("Ben");
            awarded.Milestone1Presented = new DateTime(2026, 3, 5);
            awarded.Milestone2Awarded = new DateTime(2026, 4, 6);
            var withDates = RowOf(awarded);

            Assert.AreEqual(40, bare.Cells.Sum(c => c.Span));
            Assert.AreEqual(40, withDates.Cells.Sum(c => c.Span)); // a milestone that is one date still covers its five columns
        }

        [TestMethod]
        public void ScoutsAreByNameOnOneSheetOrOnASheetForEachPatrol()
        {
            var scouts = new[] { Scout("Zoe", "Red"), Scout("Amy", "Red"), Scout("Ben", "Blue"), Scout("Pat", "") };

            var unit = WallchartLayout.Build(scouts, false).Single();
            Assert.AreEqual("", unit.Patrol);
            CollectionAssert.AreEqual(new[] { "Amy", "Ben", "Pat", "Zoe" }, unit.Rows.Select(r => r.Name).ToList());

            var patrols = WallchartLayout.Build(scouts, true);
            CollectionAssert.AreEqual(new[] { "", "Blue", "Red" }, patrols.Select(p => p.Patrol).ToList());
            CollectionAssert.AreEqual(new[] { "Amy", "Zoe" }, patrols.Single(p => p.Patrol == "Red").Rows.Select(r => r.Name).ToList());
        }

        [TestMethod]
        public void NobodyGivesNoSheets()
        {
            Assert.AreEqual(0, WallchartLayout.Build(new List<WallchartItemModel>(), false).Count);
            Assert.AreEqual(0, WallchartLayout.Build(new List<WallchartItemModel>(), true).Count);
        }

        [TestMethod]
        public void AMilestoneWholeCountIsGreenAndAnythingLessIsLeftUnshaded()
        {
            var scout = Scout("Amy");
            scout.Milestone1Community = 6; scout.Milestone1Outdoors = 5; scout.Milestone1Lead = 2;
            scout.Milestone2Community = 5; scout.Milestone2Outdoors = 6; scout.Milestone2Lead = 3;
            scout.Milestone3Community = 4; scout.Milestone3Outdoors = 9; scout.Milestone3Creative = 3; scout.Milestone3Lead = 6;
            var row = RowOf(scout);

            // milestone 1 and 2 targets are exact; milestone 3 can go past its target
            Assert.AreEqual(WallchartLayout.Complete, At(row, "Community (6)").Colour);
            Assert.IsNull(At(row, "Outdoors (6)").Colour);
            Assert.AreEqual(WallchartLayout.Complete, At(row, "Community (5)").Colour);
            Assert.IsNull(At(row, "Outdoors (5)").Colour); // six is not the five this milestone asks for
            Assert.AreEqual(WallchartLayout.Complete, At(row, "Community (4)").Colour);
            Assert.AreEqual(WallchartLayout.Complete, At(row, "Outdoors (4)").Colour);
            Assert.IsNull(At(row, "Creative (4)").Colour);

            // lead counts are green when met and grey otherwise
            Assert.AreEqual(WallchartLayout.Complete, At(row, "Lead (2)").Colour);
            Assert.AreEqual(WallchartLayout.Grey, At(row, "Lead (4)").Colour);
            Assert.AreEqual(WallchartLayout.Complete, At(row, "Lead (6)").Colour);
        }

        [TestMethod]
        public void APresentedMilestoneIsOneDateAndAnAwardedOneIsStarredUntilItIsPresented()
        {
            var scout = Scout("Amy");
            scout.Milestone1Awarded = new DateTime(2026, 3, 5);
            scout.Milestone1Presented = new DateTime(2026, 3, 20);
            scout.Milestone2Awarded = new DateTime(2026, 4, 6);
            scout.Milestone1Community = 6; // the counts are not shown once there is a date
            var row = RowOf(scout);

            var presented = At(row, "Community (6)");
            Assert.AreEqual(WallchartCellKind.PresentedDate, presented.Kind);
            Assert.AreEqual(new DateTime(2026, 3, 20), presented.Date);
            Assert.AreEqual(5, presented.Span);
            Assert.AreEqual(WallchartLayout.Complete, presented.Colour);

            var awarded = At(row, "Community (5)");
            Assert.AreEqual(WallchartCellKind.AwardedDate, awarded.Kind);
            Assert.AreEqual(new DateTime(2026, 4, 6), awarded.Date);
            Assert.AreEqual(5, awarded.Span);

            Assert.AreEqual(WallchartCellKind.Count, At(row, "Community (4)").Kind); // not awarded yet, so still counts
        }

        [TestMethod]
        public void TheIntroDatesAreShadedEvenWhenBlankAndYearsOfServiceAlwaysShows()
        {
            var scout = Scout("Amy");
            scout.IntroToScouting = new DateTime(2021, 4, 26);
            var row = RowOf(scout);

            Assert.AreEqual(WallchartCellKind.Date, At(row, "Intro To Scouting").Kind);
            Assert.AreEqual(WallchartCellKind.Blank, At(row, "Intro To Section").Kind);
            Assert.AreEqual(WallchartLayout.Grey, At(row, "Intro To Section").Colour);
            Assert.AreEqual(WallchartCellKind.Number, At(row, "Years of Service").Kind);
            Assert.AreEqual(0, At(row, "Years of Service").Count); // unknown years show as nothing yet, not blank
            Assert.IsNull(At(row, "Years of Service").Colour);
        }

        [TestMethod]
        public void ColumnsAreShadedByTheirGroupAndTheTopAwardShowsPercentUntilAwarded()
        {
            var row = RowOf(Scout("Amy"));
            var colours = new[] { "Bushcraft", "Alpine", "Aquatics", "Total Progressions", "Adventure & Sport", "Arts & Literature", "Leadership Course", "Adventurous Journey", "Personal Reflection" }
                .Select(c => At(row, c).Colour).ToArray();
            CollectionAssert.AreEqual(new[] { "#8fbc8b", "#ffa07a", "#add8e6", "#dcdcdc", "#ffe4e1", "#ffffe0", "#add8e6", "#8fbc8b", "#ffa07a" }, colours);

            var peak = At(row, "PEAK AWARD");
            Assert.AreEqual(WallchartCellKind.Percent, peak.Kind);
            Assert.AreEqual("#d8bfd8", peak.Colour);

            var awarded = Scout("Ben");
            awarded.PeakAward = 1;
            awarded.PeakAwardAwarded = new DateTime(2026, 6, 1);
            var awardedPeak = At(RowOf(awarded), "PEAK AWARD");
            Assert.AreEqual(WallchartCellKind.Date, awardedPeak.Kind);
            Assert.AreEqual(new DateTime(2026, 6, 1), awardedPeak.Date);
        }
    }
}
