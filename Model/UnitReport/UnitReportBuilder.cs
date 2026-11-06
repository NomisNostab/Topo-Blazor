using Topo.Model.History;
using Topo.Model.Members;
using Topo.Model.Milestone;
using Topo.Model.Program;
using Topo.Model.Progress;
using AdditionalAwardResults = Topo.Model.AdditionalAwards.GetUnitAchievementsResultModel;
using OasResults = Topo.Model.OAS.GetUnitAchievementsResultsModel;
using SiaResults = Topo.Model.SIA.GetSIAResultsModel;

namespace Topo.Model.UnitReport
{
    // One member's badge records, as Terrain returns them for that member
    public class UnitMemberBadges
    {
        public GetMilestoneResultModel Milestones { get; set; } = new GetMilestoneResultModel();
        public OasResults Oas { get; set; } = new OasResults();
        public SiaResults Sia { get; set; } = new SiaResults();
        public GetPeakAwardsResultsModel Peak { get; set; } = new GetPeakAwardsResultsModel();
    }

    public class UnitEventInput
    {
        // group, unit, patrol or member (a project)
        public string Scope { get; set; } = string.Empty;
        public GetEventResultModel Event { get; set; } = new GetEventResultModel();
    }

    public class UnitReportInput
    {
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        public DateTime AsOf { get; set; } = DateTime.Now;
        public List<MemberListModel> CurrentMembers { get; set; } = new List<MemberListModel>();
        public Dictionary<string, UnitMemberBadges> BadgesByMember { get; set; } = new Dictionary<string, UnitMemberBadges>();
        public AdditionalAwardResults AdditionalAwards { get; set; } = new AdditionalAwardResults();
        public Dictionary<string, string> AwardNames { get; set; } = new Dictionary<string, string>();
        public List<UnitEventInput> Events { get; set; } = new List<UnitEventInput>();
    }

    public static class UnitReportBuilder
    {
        public static UnitReportModel Build(UnitReportInput input)
        {
            var model = new UnitReportModel
            {
                UnitName = input.UnitName,
                Section = input.Section,
                From = input.From,
                To = input.To
            };

            // keyed by member number; a person listed twice is one person
            var current = input.CurrentMembers
                .GroupBy(m => PersonKey(m.member_number, m.id), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
            model.AsAt = input.AsOf;
            model.Members.CurrentLeaders = LeadersOf(current.Values).Count;
            model.Members.YouthHelpers = current.Values.Count(IsYouthHelper);
            model.Members.CurrentYouth = current.Values.Count(m => !IsLeader(m));

            AddEvents(input, model, current);
            AddBadges(input, model);
            return model;
        }

        private static void AddEvents(UnitReportInput input, UnitReportModel model, Dictionary<string, MemberListModel> current)
        {
            var former = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seen = new HashSet<string>();
            foreach (var item in input.Events)
            {
                var evt = item.Event;
                if (!seen.Add(evt.id))
                    continue;
                // only events that have happened in the range, and not cancelled ones
                if (evt.start_datetime.Date < input.From.Date || evt.start_datetime.Date > input.To.Date || evt.start_datetime > input.AsOf)
                    continue;
                if (evt.status.Contains("cancel", StringComparison.OrdinalIgnoreCase))
                    continue;

                model.EventsCounted++;

                // Everyone recorded as attending. A group event also lists other units' members, so for those only
                // the unit's own members are counted.
                var attendance = evt.attendance;
                var leaderKeys = (attendance?.leader_members ?? new Attendance_Members[0])
                    .Select(a => PersonKey(a.member_number, a.id)).ToHashSet(StringComparer.OrdinalIgnoreCase);
                var attendees = (attendance?.attendee_members ?? new Attendance_Members[0])
                    .Concat(attendance?.participant_members ?? new Attendance_Members[0])
                    .Concat(attendance?.assistant_members ?? new Attendance_Members[0])
                    .Concat(attendance?.leader_members ?? new Attendance_Members[0])
                    .GroupBy(a => PersonKey(a.member_number, a.id), StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.Key)
                    .Where(key => item.Scope != "group" || current.ContainsKey(key))
                    .ToList();

                var leaders = 0;
                var youth = 0;
                foreach (var key in attendees)
                {
                    var isLeader = current.TryGetValue(key, out var member) ? IsLeader(member) : leaderKeys.Contains(key);
                    if (isLeader)
                        leaders++;
                    else
                        youth++;
                    if (!current.ContainsKey(key))
                        former.Add(key);
                }

                if (attendees.Count == 0)
                    model.EventsWithoutAttendance++;

                var hours = Math.Max(0, (evt.end_datetime - evt.start_datetime).TotalHours);
                model.LeaderHours += hours * leaders;
                model.YouthHours += hours * youth;

                var nights = Math.Max(0, (evt.end_datetime.Date - evt.start_datetime.Date).Days);
                model.NightsRun += nights;
                model.YouthPersonNights += nights * youth;
                model.LeaderPersonNights += nights * leaders;
                model.PersonNightsCamped += nights * (youth + leaders);

                var logbook = evt.achievement_pathway_logbook_data;
                if (logbook != null && (logbook.categories ?? new object[0]).Any(c => string.Equals(c?.ToString(), "walking_hike", StringComparison.OrdinalIgnoreCase)))
                {
                    model.HikeEvents++;
                    var km = logbook.distance_walkabout / 1000.0;
                    if (km > 0)
                        model.HikeEventsWithDistance++;
                    model.HikeKmOnce += km;
                    model.HikeKmPerPerson += km * (youth + leaders);
                }
            }

            model.Members.FormerAttendees = former.Count;
            model.LeaderHours = Math.Round(model.LeaderHours, 1);
            model.YouthHours = Math.Round(model.YouthHours, 1);
            model.HikeKmOnce = Math.Round(model.HikeKmOnce, 1);
            model.HikeKmPerPerson = Math.Round(model.HikeKmPerPerson, 1);
        }

        // Badges awarded to the unit's current youth in the range, while in this unit's section
        private static void AddBadges(UnitReportInput input, UnitReportModel model)
        {
            var section = (input.Section ?? "").Trim().ToLowerInvariant();
            bool InUnit(string recordSection) => (recordSection ?? "").Trim().ToLowerInvariant() == section;
            bool InRange(string imported, DateTime updated)
            {
                var date = ScoutingHistoryBuilder.AwardedDate(imported, updated);
                return date >= input.From.Date && date <= input.To.Date;
            }

            var youthIds = input.CurrentMembers.Where(m => !IsLeader(m)).Select(m => m.id).ToHashSet();
            var milestones = new SortedDictionary<int, int>();
            var oas = new SortedDictionary<string, int>();
            var sia = new SortedDictionary<string, int>();
            foreach (var badges in input.BadgesByMember.Where(b => youthIds.Contains(b.Key)).Select(b => b.Value))
            {
                foreach (var r in badges.Milestones.results.Where(r => r.status == "awarded" && InUnit(r.section) && InRange(r.imported.date_awarded, r.status_updated)))
                    milestones[r.achievement_meta.stage] = milestones.GetValueOrDefault(r.achievement_meta.stage) + 1;

                foreach (var r in badges.Oas.results.Where(r => (r.status == "awarded" || r.status == "approved") && InUnit(r.section) && InRange(r.imported.date_awarded, r.status_updated)))
                {
                    var stream = ScoutingHistoryBuilder.Title(r.achievement_meta.stream);
                    oas[stream] = oas.GetValueOrDefault(stream) + (r.achievement_meta.stage >= 7 ? 2 : 1);
                }

                foreach (var r in badges.Sia.results.Where(r => r.status == "awarded" && InUnit(r.section) && InRange(r.imported.date_awarded, r.status_updated)))
                {
                    var area = ScoutingHistoryBuilder.SiaName(r.answers.special_interest_area_selection ?? "");
                    area = area == "" ? "Other" : area;
                    sia[area] = sia.GetValueOrDefault(area) + 1;
                }

                model.Badges.PeakAwards += badges.Peak.results.Count(r => r.status == "awarded" && InUnit(r.section) && InRange("", r.status_updated));
            }

            var awards = new SortedDictionary<string, int>();
            foreach (var r in input.AdditionalAwards.results.Where(r => r.status == "awarded" && youthIds.Contains(r.member_id) && InUnit(r.section) && InRange(r.imported.date_awarded, r.status_updated)))
            {
                var name = input.AwardNames.TryGetValue(r.achievement_meta.additional_award_id, out var title)
                    ? title
                    : ScoutingHistoryBuilder.Title(r.achievement_meta.additional_award_id);
                awards[name] = awards.GetValueOrDefault(name) + 1;
            }

            model.Badges.Milestones = milestones.Select(m => new UnitBadgeCount { Name = $"Milestone {m.Key}", Count = m.Value }).ToList();
            model.Badges.OasStreams = oas.Select(o => new UnitBadgeCount { Name = o.Key, Count = o.Value }).ToList();
            model.Badges.SiaCategories = sia.Select(s => new UnitBadgeCount { Name = s.Key, Count = s.Value }).ToList();
            model.Badges.AdditionalAwards = awards.Select(a => new UnitBadgeCount { Name = a.Key, Count = a.Value }).ToList();
        }

        // Adult leaders only. A unit leader (UL) is a youth member, who chairs the unit council, so counts with the youth.
        public static bool IsLeader(MemberListModel member) => member.isAdultLeader == 1;

        // Terrain gives the adult leader duty to others in a unit too: young helpers, and people whose role is in the group
        // (a group leader or advisor). A helper is under 18. The unit's own leaders are the adults in its "Adult Leaders"
        // patrol; when nobody has a patrol that can't be told, so every adult counts.
        public static List<MemberListModel> LeadersOf(IEnumerable<MemberListModel> members)
        {
            var adults = members.Where(m => IsLeader(m) && !IsUnder18(m)).ToList();
            var inAPatrol = adults.Where(m => !string.IsNullOrEmpty(m.patrol_name)).ToList();
            return inAPatrol.Count > 0 ? inAPatrol : adults;
        }

        public static bool IsYouthHelper(MemberListModel member) => IsLeader(member) && IsUnder18(member);

        private static bool IsUnder18(MemberListModel member) => int.TryParse((member.age ?? "").Split('y')[0], out var years) && years < 18;

        private static string PersonKey(string memberNumber, string id) => string.IsNullOrEmpty(memberNumber) ? id : memberNumber;
    }
}
