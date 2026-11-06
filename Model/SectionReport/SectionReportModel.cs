using Topo.Model.UnitReport;

namespace Topo.Model.SectionReport
{
    // Why someone is no longer in the unit. Terrain only says who is a member now, so the leader picks one for
    // anyone Topo can't place itself.
    public enum DepartureStatus
    {
        Unset,
        MovedUp,
        MovedGroup,
        LeftScouting
    }

    public class DepartedMember
    {
        public string Id { get; set; } = string.Empty;
        public string MemberNumber { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        // Where Topo found them, e.g. "Now in Scouts"
        public string Source { get; set; } = string.Empty;
        public DepartureStatus Status { get; set; } = DepartureStatus.Unset;
    }

    // Where a new member came from; it decides which investiture items they are given
    public enum NewMemberOrigin
    {
        // New to Scouting
        OffTheStreet,
        // Linked up from a younger section of this group
        YoungerSection,
        // Joined from another group
        AnotherGroup
    }

    public class NewMember
    {
        public string Name { get; set; } = string.Empty;
        public NewMemberOrigin Origin { get; set; }
    }

    public class LeaderEntry
    {
        public string Name { get; set; } = string.Empty;
        public string MemberNumber { get; set; } = string.Empty;
    }

    public class AgeBucket
    {
        // Whole years
        public int Age { get; set; }
        public int Count { get; set; }
    }

    // How old the unit's youth are and how evenly they are spread over the section's age range
    public class AgeProfile
    {
        // The youth whose age is known
        public int Counted { get; set; }
        public double AverageYears { get; set; }
        public double YoungestYears { get; set; }
        public double OldestYears { get; set; }
        // The ages of the section, widened to take in anyone outside them
        public List<AgeBucket> Buckets { get; set; } = new List<AgeBucket>();
        // 0 to 1: 1 means the same number of youth at every age of the range. Null when there are too few to say.
        public double? Evenness { get; set; }
        // Where the average sits in the range: -1 is the young end, 1 the old end
        public double Lean { get; set; }

        public static string Format(double years) => $"{(int)years}y {(int)((years - (int)years) * 12)}m";

        public string Spread
        {
            get
            {
                if (Evenness == null)
                    return "Too few members to say";
                var how = Evenness >= 0.85 ? "Evenly spread" : Evenness >= 0.65 ? "Fairly even" : "Bunched together";
                var lean = Lean < -0.25 ? ", leaning younger" : Lean > 0.25 ? ", leaning older" : "";
                return $"{how}{lean} ({Evenness:P0})";
            }
        }
    }

    public class ReportEvent
    {
        public string Title { get; set; } = string.Empty;
        public DateTime Start { get; set; }
        public string ChallengeArea { get; set; } = string.Empty;
        public int YouthAttended { get; set; }
        public int LeadersAttended { get; set; }
    }

    public class OasStageCount
    {
        public string Stream { get; set; } = string.Empty;
        public int Stage { get; set; }
        public int Count { get; set; }
    }

    public class SectionBadges
    {
        public int PeakAwards { get; set; }
        public List<UnitBadgeCount> Milestones { get; set; } = new List<UnitBadgeCount>();
        public int MilestonesTotal => Milestones.Sum(m => m.Count);
        // The things a new member is given when they join (tape, scarf, badges), counted from the new members
        public List<UnitBadgeCount> InvestitureItems { get; set; } = new List<UnitBadgeCount>();
        // Every Outdoor Adventure Skill stage awarded in the period; a stage of 7 or more counts as two progressions
        public int OasProgressions { get; set; }
        // Badges per stream and stage. A member who jumps several levels at once only needs the highest, so it is one
        // badge per member, stream and day awarded.
        public List<OasStageCount> OasStages { get; set; } = new List<OasStageCount>();
        public int OasBadgesTotal => OasStages.Sum(o => o.Count);
        public List<UnitBadgeCount> SiaCategories { get; set; } = new List<UnitBadgeCount>();
        public int SiaTotal => SiaCategories.Sum(s => s.Count);
        public List<UnitBadgeCount> AdditionalAwards { get; set; } = new List<UnitBadgeCount>();
        public int AdditionalAwardsTotal => AdditionalAwards.Sum(a => a.Count);
    }

    // One unit's figures for the period
    public class SectionReportModel
    {
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        // The section a member moves up to, if there is one (Joeys to Cubs, and so on)
        public string NextSection { get; set; } = string.Empty;

        public int CurrentYouth { get; set; }
        // Invested = has an Introduction to Section for this section
        public int InvestedYouth { get; set; }
        public int InvestedInPeriod { get; set; }

        public List<NewMember> NewMembers { get; set; } = new List<NewMember>();
        public int NewFromYoungerSection => NewMembers.Count(n => n.Origin == NewMemberOrigin.YoungerSection);
        public int NewOffTheStreet => NewMembers.Count(n => n.Origin == NewMemberOrigin.OffTheStreet);
        public int NewFromAnotherGroup => NewMembers.Count(n => n.Origin == NewMemberOrigin.AnotherGroup);

        public List<DepartedMember> Departed { get; set; } = new List<DepartedMember>();
        public int MovedUp => Departed.Count(d => d.Status == DepartureStatus.MovedUp);
        public int MovedGroup => Departed.Count(d => d.Status == DepartureStatus.MovedGroup);
        public int LeftScouting => Departed.Count(d => d.Status == DepartureStatus.LeftScouting);
        public int DepartedUnset => Departed.Count(d => d.Status == DepartureStatus.Unset);

        public int CurrentLeaders { get; set; }
        // Who the current leaders are; only used to count each person once across a group
        public List<LeaderEntry> Leaders { get; set; } = new List<LeaderEntry>();
        // Young people with a leader's duty in the unit, who are not counted as youth or leaders
        public List<LeaderEntry> YouthHelpers { get; set; } = new List<LeaderEntry>();
        public AgeProfile Ages { get; set; } = new AgeProfile();
        public List<string> NewLeaders { get; set; } = new List<string>();
        public List<string> DepartedLeaders { get; set; } = new List<string>();

        // When a roster was saved by an earlier report; new and departed leaders can only be named against one
        public DateTime? SavedRosterDate { get; set; }

        public List<ReportEvent> PastEvents { get; set; } = new List<ReportEvent>();
        public List<ReportEvent> UpcomingEvents { get; set; } = new List<ReportEvent>();

        public SectionBadges Badges { get; set; } = new SectionBadges();
    }

    // The report for one unit, or for every unit the person can see (the group), with the totals
    public class SectionReportSet
    {
        public DateTime GeneratedAt { get; set; }
        public string GeneratedBy { get; set; } = string.Empty;
        public string GroupName { get; set; } = string.Empty;
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        public bool IsGroup { get; set; }
        public List<SectionReportModel> Sections { get; set; } = new List<SectionReportModel>();

        public int TotalYouth => Sections.Sum(s => s.CurrentYouth);
        public int TotalInvested => Sections.Sum(s => s.InvestedYouth);
        public int TotalInvestedInPeriod => Sections.Sum(s => s.InvestedInPeriod);
        public int TotalNewFromYounger => Sections.Sum(s => s.NewFromYoungerSection);
        public int TotalNewOffTheStreet => Sections.Sum(s => s.NewOffTheStreet);
        public int TotalNewFromAnotherGroup => Sections.Sum(s => s.NewFromAnotherGroup);
        // Investiture items across every unit, each item once
        public List<UnitBadgeCount> TotalInvestitureItems => Sections
            .SelectMany(s => s.Badges.InvestitureItems)
            .GroupBy(i => i.Name)
            .Select(g => new UnitBadgeCount { Name = g.Key, Count = g.Sum(i => i.Count) })
            .ToList();
        public int TotalMovedUp => Sections.Sum(s => s.MovedUp);
        public int TotalMovedGroup => Sections.Sum(s => s.MovedGroup);
        public int TotalLeftScouting => Sections.Sum(s => s.LeftScouting);
        // An adult who leads in two sections is one leader
        public int TotalLeaders => Sections.SelectMany(s => s.Leaders)
            .Select(l => string.IsNullOrEmpty(l.MemberNumber) ? l.Name : l.MemberNumber)
            .Distinct(StringComparer.OrdinalIgnoreCase).Count();
        public int TotalYouthHelpers => Sections.SelectMany(s => s.YouthHelpers)
            .Select(l => string.IsNullOrEmpty(l.MemberNumber) ? l.Name : l.MemberNumber)
            .Distinct(StringComparer.OrdinalIgnoreCase).Count();
        public int TotalPeakAwards => Sections.Sum(s => s.Badges.PeakAwards);
        public int TotalMilestones => Sections.Sum(s => s.Badges.MilestonesTotal);
        public int TotalOasBadges => Sections.Sum(s => s.Badges.OasBadgesTotal);
        public int TotalOasProgressions => Sections.Sum(s => s.Badges.OasProgressions);
        public int TotalSia => Sections.Sum(s => s.Badges.SiaTotal);
        public int TotalAdditionalAwards => Sections.Sum(s => s.Badges.AdditionalAwardsTotal);
    }
}
