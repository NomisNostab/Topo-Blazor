using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Topo.Model.Members;

namespace Topo.Model.Progress
{
    public class ProgressDetailsPageViewModel
    {
        public MemberListModel Member { get; set; } = new MemberListModel();
        public string GroupName { get; set; } = string.Empty;
        public string IntroToScoutingDate { get; set; } = string.Empty;
        public string IntroToSectionDate { get; set; } = string.Empty;
        public List<MilestoneSummary> Milestones { get; set; } = new List<MilestoneSummary>();
        public List<OASSummary> OASSummaries { get; set; } = new List<OASSummary>();
        public Stats Stats { get; set; } = new Stats();
        public List<SIASummary> SIASummaries { get; set; } = new List<SIASummary>();
        public PeakAward PeakAward { get; set; } = new PeakAward();
        public string? DisableOAS { get; set; }
        public string? DisableCoreOAS { get; set; }
    }

    public class MilestoneSummary
    {
        public int Milestone { get; set; }
        public string Status { get; set; } = string.Empty;
        public DateTime Awarded { get; set; }
        public List<MilestoneLog> ParticipateLogs { get; set; } = new List<MilestoneLog>();
        public List<MilestoneLog> AssistLogs { get; set; } = new List<MilestoneLog>();
        public List<MilestoneLog> LeadLogs { get; set; } = new List<MilestoneLog>();
    }

    public class MilestoneLog
    {
        public string ChallengeArea { get; set; } = string.Empty;
        public string EventName { get; set; } = string.Empty;
        public DateTime EventDate { get; set; }
    }

    public class OASSummary
    {
        public string Stream { get; set; } = string.Empty;
        public int Stage { get; set; }
        public DateTime Awarded { get; set; }
        public string Section { get; set; } = string.Empty;
        public string Template { get; set; } = string.Empty;
    }

    public class Stats
    {
        public int OasProgressions { get; set; }
        public int KmsHiked { get; set; }
        public int NightsCamped { get; set; }
        public int NightsCampedInSection { get; set; }
    }

    public class SIASummary
    {
        public string Area { get; set; } = string.Empty;
        public string Project { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public DateTime Awarded { get; set; }
    }

    public class PeakAward
    {
        public string PersonalDevelopmentCourse { get; set; } = string.Empty;
        public string AdventurousJourney { get; set; } = string.Empty;
        public string PersonalReflection { get; set; } = string.Empty;
        public string Awarded { get; set; } = string.Empty;
    }
}
