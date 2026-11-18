namespace Topo.Model.Wallchart
{
    public enum WallchartCellKind
    {
        // Nothing to show, though the cell may still be shaded
        Blank,
        // A count, left empty when it is nothing
        Count,
        // A number that is shown even when it is nothing
        Number,
        // A date, in the short form the columns are sized for
        Date,
        // A milestone's presented date, which stands for the whole milestone and spans its columns
        PresentedDate,
        // A milestone's awarded date, starred because it has not been presented yet
        AwardedDate,
        // A fraction of the way to an award
        Percent
    }

    public class WallchartCell
    {
        public WallchartCellKind Kind { get; set; }
        public int Count { get; set; }
        public DateTime? Date { get; set; }
        public double Fraction { get; set; }
        // A hex colour, or null to leave the cell unshaded
        public string? Colour { get; set; }
        // How many columns the cell covers
        public int Span { get; set; } = 1;
    }

    public class WallchartRow
    {
        public string Name { get; set; } = string.Empty;
        // The columns after the name, in the order of WallchartLayout.Columns; a spanning cell stands for several
        public List<WallchartCell> Cells { get; set; } = new List<WallchartCell>();
    }

    // A sheet of the wallchart: the whole unit, or one patrol when the report is broken up by patrol
    public class WallchartSheet
    {
        // Blank for the whole unit
        public string Patrol { get; set; } = string.Empty;
        public List<WallchartRow> Rows { get; set; } = new List<WallchartRow>();
    }

    public static class WallchartLayout
    {
        // The shades the columns are given
        public const string Complete = "#9cf250";
        public const string Grey = "#dcdcdc";
        private const string Green = "#8fbc8b";
        private const string Salmon = "#ffa07a";
        private const string Blue = "#add8e6";
        private const string Pink = "#ffe4e1";
        private const string Yellow = "#ffffe0";
        private const string Purple = "#d8bfd8";

        // The headings over the columns after the name, and how many columns each covers
        public static readonly (string Heading, int Span)[] Groups =
        {
            ("", 3), ("MILESTONE 1", 5), ("MILESTONE 2", 5), ("MILESTONE 3", 5), ("OUTDOOR ADVENTURE SKILLS", 12), ("SPECIAL INTEREST AREAS", 6), ("", 4)
        };

        // The headings over the Outdoor Adventure Skill columns
        public static readonly (string Heading, int Span)[] OutdoorGroups = { ("CORE", 3), ("LAND", 3), ("WATER", 3) };

        public static readonly string[] Columns =
        {
            "Intro To Scouting", "Intro To Section", "Years of Service",
            "Community (6)", "Outdoors (6)", "Creative (6)", "Personal Growth (6)", "Lead (2)",
            "Community (5)", "Outdoors (5)", "Creative (5)", "Personal Growth (5)", "Lead (4)",
            "Community (4)", "Outdoors (4)", "Creative (4)", "Personal Growth (4)", "Lead (6)",
            "Bushcraft", "Bushwalking", "Camping", "Alpine", "Cycling", "Vertical", "Aquatics", "Boating", "Paddling",
            "Total Progressions", "Total Nights Camped", "Total KMs Hiked",
            "Adventure & Sport", "Arts & Literature", "Environment", "STEM & Innovation", "Growth & Development", "Creating a Better World",
            "Leadership Course", "Adventurous Journey", "Personal Reflection", "PEAK AWARD"
        };

        // How the Wallchart is laid out, shared by the page's preview and the pdf and spreadsheet
        public static List<WallchartSheet> Build(IEnumerable<WallchartItemModel> items, bool breakByPatrol) =>
            items
                .GroupBy(i => breakByPatrol ? i.MemberPatrol : "")
                .OrderBy(g => g.Key)
                .Select(patrol => new WallchartSheet
                {
                    Patrol = patrol.Key ?? "",
                    Rows = patrol.OrderBy(i => i.MemberName).Select(Row).ToList()
                })
                .ToList();

        private static WallchartRow Row(WallchartItemModel w)
        {
            var cells = new List<WallchartCell>
            {
                DateCell(w.IntroToScouting, Grey),
                DateCell(w.IntroToSection, Grey),
                new WallchartCell { Kind = WallchartCellKind.Number, Count = w.YearsOfService ?? 0 }
            };
            cells.AddRange(Milestone(1, w.Milestone1Presented, w.Milestone1Awarded, w.Milestone1Community, w.Milestone1Outdoors, w.Milestone1Creative, w.Milestone1PersonalGrowth, w.Milestone1Lead));
            cells.AddRange(Milestone(2, w.Milestone2Presented, w.Milestone2Awarded, w.Milestone2Community, w.Milestone2Outdoors, w.Milestone2Creative, w.Milestone2PersonalGrowth, w.Milestone2Lead));
            cells.AddRange(Milestone(3, w.Milestone3Presented, w.Milestone3Awarded, w.Milestone3Community, w.Milestone3Outdoors, w.Milestone3Creative, w.Milestone3PersonalGrowth, w.Milestone3Lead));
            cells.AddRange(new[]
            {
                CountCell(w.OASBushcraftStage, Green), CountCell(w.OASBushwalkingStage, Green), CountCell(w.OASCampingStage, Green),
                CountCell(w.OASAlpineStage, Salmon), CountCell(w.OASCyclingStage, Salmon), CountCell(w.OASVerticalStage, Salmon),
                CountCell(w.OASAquaticsStage, Blue), CountCell(w.OASBoatingStage, Blue), CountCell(w.OASPaddlingStage, Blue),
                CountCell(w.OASStageProgressions, Grey), CountCell(w.NightsCamped, Grey), CountCell(w.KMsHiked, Grey),
                CountCell(w.SIAAdventureSport, Pink), CountCell(w.SIAArtsLiterature, Yellow), CountCell(w.SIAEnvironment, Pink),
                CountCell(w.SIAStemInnovation, Yellow), CountCell(w.SIAGrowthDevelopment, Pink), CountCell(w.SIACreatingABetterWorld, Yellow),
                DateCell(w.LeadershipCourse, Blue), DateCell(w.AdventurousJourney, Green), DateCell(w.PersonalReflection, Salmon),
                // the percentage towards the top award until it is granted, then the day it was
                w.PeakAwardAwarded.HasValue
                    ? DateCell(w.PeakAwardAwarded, Purple)
                    : new WallchartCell { Kind = WallchartCellKind.Percent, Fraction = w.PeakAward, Colour = Purple }
            });
            return new WallchartRow { Name = w.MemberName, Cells = cells };
        }

        // A milestone is one date once it has been presented or awarded; until then its four participate counts and its lead count
        private static IEnumerable<WallchartCell> Milestone(int level, DateTime? presented, DateTime? awarded, int community, int outdoors, int creative, int personalGrowth, int lead)
        {
            if (presented.HasValue)
                return new[] { new WallchartCell { Kind = WallchartCellKind.PresentedDate, Date = presented, Colour = Complete, Span = 5 } };
            if (awarded.HasValue)
                return new[] { new WallchartCell { Kind = WallchartCellKind.AwardedDate, Date = awarded, Colour = Complete, Span = 5 } };

            // the count that is the whole of what the milestone asks for is shaded green
            var participateTarget = level == 1 ? 6 : level == 2 ? 5 : 4;
            var leadTarget = level == 1 ? 2 : level == 2 ? 4 : 6;
            // milestone 3 lets a count go past its target, as the milestone 1 and 2 targets are exact
            bool Met(int count, int target) => level == 3 ? count >= target : count == target;
            return new[] { community, outdoors, creative, personalGrowth }
                .Select(c => new WallchartCell { Kind = WallchartCellKind.Count, Count = c, Colour = Met(c, participateTarget) ? Complete : null })
                .Append(new WallchartCell { Kind = WallchartCellKind.Count, Count = lead, Colour = Met(lead, leadTarget) ? Complete : Grey });
        }

        private static WallchartCell CountCell(int count, string colour) => new WallchartCell { Kind = WallchartCellKind.Count, Count = count, Colour = colour };

        private static WallchartCell DateCell(DateTime? date, string colour) => new WallchartCell { Kind = date.HasValue ? WallchartCellKind.Date : WallchartCellKind.Blank, Date = date, Colour = colour };
    }
}
