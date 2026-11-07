namespace Topo.Model.Milestone
{
    // A count in a section's table and the shade it is given
    public class MilestoneCell
    {
        public int Count { get; set; }
        public string Colour { get; set; } = MilestoneColours.White;
    }

    public class MilestoneRow
    {
        public string Name { get; set; } = string.Empty;
        public int PercentComplete { get; set; }
        // Community, Outdoors, Creative, Personal Growth
        public List<MilestoneCell> Participate { get; set; } = new List<MilestoneCell>();
        public MilestoneCell Lead { get; set; } = new MilestoneCell();
        // Which challenge areas the lead (and assist) roles were in, such as "Co-2 O"
        public string LeadAreas { get; set; } = string.Empty;
    }

    // Everyone at one milestone: 1, 2 or 3, or 4 for those who have completed Milestone 3
    public class MilestoneSection
    {
        public int Key { get; set; }
        public string Title { get; set; } = string.Empty;
        public string ParticipateHeading { get; set; } = string.Empty;
        public string LeadHeading { get; set; } = string.Empty;
        public List<MilestoneRow> Rows { get; set; } = new List<MilestoneRow>();

        // The average of each count column (the four participate columns, then lead), as the report's Average row shows
        public List<double> Averages() => Enumerable.Range(0, 5)
            .Select(i => Rows.Count == 0 ? 0 : Rows.Average(r => i < 4 ? r.Participate[i].Count : r.Lead.Count))
            .ToList();
    }

    // A sheet of the report: the whole unit, or one patrol when the report is broken up by patrol
    public class MilestonePatrolGroup
    {
        // Blank for the whole unit
        public string Patrol { get; set; } = string.Empty;
        public List<MilestoneSection> Sections { get; set; } = new List<MilestoneSection>();
    }

    public static class MilestoneLayout
    {
        // How the Milestone report is laid out, shared by the page's preview and the pdf and spreadsheet
        public static List<MilestonePatrolGroup> Build(IEnumerable<MilestoneSummaryListModel> summaries, bool breakByPatrol) =>
            summaries
                .GroupBy(s => breakByPatrol ? s.memberPatrol : "")
                .OrderBy(g => g.Key)
                .Select(patrol => new MilestonePatrolGroup
                {
                    Patrol = patrol.Key ?? "",
                    // there is no Milestone 4, so a Scout who has finished Milestone 3 gets a section of their own as key 4
                    Sections = patrol
                        .GroupBy(s => s.currentLevel == 3 && s.milestone3Awarded ? 4 : s.currentLevel)
                        .OrderBy(g => g.Key)
                        .Select(level => Section(level.Key, level))
                        .ToList()
                })
                .ToList();

        private static MilestoneSection Section(int key, IEnumerable<MilestoneSummaryListModel> members)
        {
            // key 4 uses Milestone 3's own counts and shading
            var level = Math.Min(key, 3);
            return new MilestoneSection
            {
                Key = key,
                Title = key == 4 ? "Milestone 3 - Completed" : $"Milestone {key}",
                ParticipateHeading = key switch { 1 => "Participate - 6", 2 => "Participate - 5", 3 or 4 => "Participate - 4", _ => "Participate" },
                LeadHeading = key switch { 1 => "Lead - 2", 2 => "Lead - 4", 3 or 4 => "Lead - 6", _ => "Lead" },
                Rows = members.OrderBy(m => m.memberName).Select(m => Row(level, m)).ToList()
            };
        }

        private static MilestoneRow Row(int level, MilestoneSummaryListModel m)
        {
            int[] participate;
            int lead;
            string areas;
            switch (level)
            {
                case 1:
                    participate = new[] { m.milestone1ParticipateCommunity, m.milestone1ParticipateOutdoors, m.milestone1ParticipateCreative, m.milestone1ParticipatePersonalGrowth };
                    lead = m.milestone1Lead;
                    areas = m.milestone1LeadArea;
                    break;
                case 2:
                    participate = new[] { m.milestone2ParticipateCommunity, m.milestone2ParticipateOutdoors, m.milestone2ParticipateCreative, m.milestone2ParticipatePersonalGrowth };
                    lead = m.milestone2Lead;
                    areas = m.milestone2LeadArea;
                    break;
                default:
                    participate = new[] { m.milestone3ParticipateCommunity, m.milestone3ParticipateOutdoors, m.milestone3ParticipateCreative, m.milestone3ParticipatePersonalGrowth };
                    lead = m.milestone3Lead;
                    areas = m.milestone3LeadArea;
                    break;
            }
            return new MilestoneRow
            {
                Name = m.memberName,
                PercentComplete = m.percentComplete,
                Participate = participate.Select(c => new MilestoneCell { Count = c, Colour = MilestoneColours.For(level, true, c) }).ToList(),
                Lead = new MilestoneCell { Count = lead, Colour = MilestoneColours.For(level, false, lead) },
                LeadAreas = areas
            };
        }
    }
}
