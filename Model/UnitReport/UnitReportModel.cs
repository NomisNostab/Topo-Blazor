namespace Topo.Model.UnitReport
{
    // Totals for one unit over a date range
    public class UnitReportModel
    {
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        public bool AllTime { get; set; }

        public int EventsCounted { get; set; }
        public int EventsWithoutAttendance { get; set; }

        public UnitMembersSummary Members { get; set; } = new UnitMembersSummary();
        public UnitBadgesSummary Badges { get; set; } = new UnitBadgesSummary();

        // Nights away
        public int NightsRun { get; set; }
        public int PersonNightsCamped { get; set; }
        public int YouthPersonNights { get; set; }
        public int LeaderPersonNights { get; set; }

        // Hiking: the distance of each hike counted once, or added up for everyone on it
        public int HikeEvents { get; set; }
        public int HikeEventsWithDistance { get; set; }
        public double HikeKmOnce { get; set; }
        public double HikeKmPerPerson { get; set; }
        public bool KmPerPerson { get; set; }
        public double HikeKm => KmPerPerson ? HikeKmPerPerson : HikeKmOnce;

        public double LeaderHours { get; set; }
        public double YouthHours { get; set; }
    }

    public class UnitMembersSummary
    {
        public int CurrentYouth { get; set; }
        public int CurrentLeaders { get; set; }
        // People who attended the unit's own events in the range but are not current members
        public int FormerAttendees { get; set; }
        public int Total => CurrentYouth + CurrentLeaders + FormerAttendees;
    }

    public class UnitBadgesSummary
    {
        public int PeakAwards { get; set; }
        public List<UnitBadgeCount> Milestones { get; set; } = new List<UnitBadgeCount>();
        public int MilestonesTotal => Milestones.Sum(m => m.Count);
        public List<UnitBadgeCount> OasStreams { get; set; } = new List<UnitBadgeCount>();
        // Stages 7 to 9 count as two progressions each
        public int OasProgressions => OasStreams.Sum(o => o.Count);
        public List<UnitBadgeCount> SiaCategories { get; set; } = new List<UnitBadgeCount>();
        public int SiaTotal => SiaCategories.Sum(s => s.Count);
        public List<UnitBadgeCount> AdditionalAwards { get; set; } = new List<UnitBadgeCount>();
        public int AdditionalAwardsTotal => AdditionalAwards.Sum(a => a.Count);
    }

    public class UnitBadgeCount
    {
        public string Name { get; set; } = string.Empty;
        public int Count { get; set; }
    }
}
