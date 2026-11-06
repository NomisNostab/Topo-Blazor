using Topo.Model.History;
using Topo.Model.Members;
using Topo.Model.Program;
using Topo.Model.Progress;
using Topo.Model.UnitReport;
using AdditionalAwardResults = Topo.Model.AdditionalAwards.GetUnitAchievementsResultModel;

namespace Topo.Model.SectionReport
{
    public class RosterPerson
    {
        public string Id { get; set; } = string.Empty;
        public string MemberNumber { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }

    // A unit's roster as it was when a report was issued. Terrain only knows who is in a unit now, so the next report
    // compares against this to name who has joined or left.
    public class RosterSnapshot
    {
        public DateTime SavedOn { get; set; }
        public List<RosterPerson> Youth { get; set; } = new List<RosterPerson>();
        public List<RosterPerson> Leaders { get; set; } = new List<RosterPerson>();
    }

    // What Terrain holds for one member that the report needs
    public class SectionMemberRecords
    {
        public UnitMemberBadges Badges { get; set; } = new UnitMemberBadges();
        // Badges are only read for members of the unit being reported on, not for the next section's members
        public bool BadgesLoaded { get; set; }
        public GetIntroductionResultsModel IntroToSection { get; set; } = new GetIntroductionResultsModel();
        public GetMemberAchievementsResultModel AllAchievements { get; set; } = new GetMemberAchievementsResultModel();
    }

    public class NextSectionMember
    {
        public RosterPerson Person { get; set; } = new RosterPerson();
        public SectionMemberRecords Records { get; set; } = new SectionMemberRecords();
    }

    public class SectionReportInput
    {
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        public DateTime AsOf { get; set; } = DateTime.Now;
        public List<MemberListModel> CurrentMembers { get; set; } = new List<MemberListModel>();
        // By member id, for the unit's current youth
        public Dictionary<string, SectionMemberRecords> Records { get; set; } = new Dictionary<string, SectionMemberRecords>();
        public AdditionalAwardResults AdditionalAwards { get; set; } = new AdditionalAwardResults();
        public Dictionary<string, string> AwardNames { get; set; } = new Dictionary<string, string>();
        public List<UnitEventInput> Events { get; set; } = new List<UnitEventInput>();
        // The ids of every unit in the group, to tell a member from this group's younger section from one who joined from
        // another group. When empty, anyone with earlier records is taken to be from this group.
        public HashSet<string> GroupUnitIds { get; set; } = new HashSet<string>();
        public RosterSnapshot? Previous { get; set; }
        // The members of the next section's unit, when the person can see it
        public List<NextSectionMember> NextSectionMembers { get; set; } = new List<NextSectionMember>();
    }

    public static class SectionReportBuilder
    {
        private static readonly string[] SectionOrder = { "joey", "cub", "scout", "venturer", "rover" };

        // What a new member is given when they join. These are the group's own practice, not something Terrain records, so
        // they are listed here to be edited if the group changes what it hands out. A member who linked up from a younger
        // section of the group is given none of these.
        private static readonly string[] OffTheStreetItems =
            { "Group Tape", "Scarf", "Scarf Badge", "Region Badge", "Australian Scouting Badge", "World Scouting Badge" };
        private static readonly string[] AnotherGroupItems =
            { "Group Tape", "Scarf", "Scarf Badge", "Region Badge" };

        public static string NextSectionOf(string section)
        {
            var index = Array.IndexOf(SectionOrder, Key(section));
            return index < 0 || index + 1 >= SectionOrder.Length ? "" : SectionOrder[index + 1];
        }

        public static string PreviousSectionOf(string section)
        {
            var index = Array.IndexOf(SectionOrder, Key(section));
            return index <= 0 ? "" : SectionOrder[index - 1];
        }

        public static RosterSnapshot TakeSnapshot(IEnumerable<MemberListModel> members, DateTime savedOn)
        {
            var list = Distinct(members);
            return new RosterSnapshot
            {
                SavedOn = savedOn,
                Youth = list.Where(m => !UnitReportBuilder.IsLeader(m)).Select(Person).ToList(),
                Leaders = LeadersOf(list).Select(Person).ToList()
            };
        }

        private static List<MemberListModel> Distinct(IEnumerable<MemberListModel> members) =>
            members.GroupBy(m => PersonKey(m.member_number, m.id), StringComparer.OrdinalIgnoreCase).Select(g => g.First()).ToList();

        public static SectionReportModel Build(SectionReportInput input)
        {
            var section = Key(input.Section);
            var model = new SectionReportModel
            {
                UnitId = input.UnitId,
                UnitName = input.UnitName,
                Section = section,
                NextSection = NextSectionOf(section),
                SavedRosterDate = input.Previous?.SavedOn
            };

            // a person listed twice is one person
            var members = Distinct(input.CurrentMembers);
            var youth = members.Where(m => !UnitReportBuilder.IsLeader(m)).ToList();
            var leaders = LeadersOf(members);
            var currentKeys = members.Select(m => PersonKey(m.member_number, m.id)).ToHashSet(StringComparer.OrdinalIgnoreCase);
            model.CurrentYouth = youth.Count;
            model.CurrentLeaders = leaders.Count;
            model.Leaders = leaders
                .Select(l => new LeaderEntry { Name = Name(l), MemberNumber = l.member_number })
                .OrderBy(l => l.Name, StringComparer.OrdinalIgnoreCase).ToList();
            model.YouthHelpers = members.Where(m => UnitReportBuilder.IsLeader(m) && IsUnder18(m))
                .Select(l => new LeaderEntry { Name = Name(l), MemberNumber = l.member_number })
                .OrderBy(l => l.Name, StringComparer.OrdinalIgnoreCase).ToList();
            model.Ages = AgeProfileOf(youth, section);

            AddInvestedAndNew(input, model, youth, section);
            AddDepartures(input, model, currentKeys, section);
            AddLeaderChanges(input, model, leaders, currentKeys);
            AddEvents(input, model);
            AddBadges(input, model, youth, section);
            return model;
        }

        // Invested = has an Introduction to Section for this section. New = invested in the period, and a new member came
        // from a younger section if Terrain holds anything for them from one, otherwise they came off the street.
        private static void AddInvestedAndNew(SectionReportInput input, SectionReportModel model, List<MemberListModel> youth, string section)
        {
            var rank = Array.IndexOf(SectionOrder, section);
            foreach (var member in youth)
            {
                if (!input.Records.TryGetValue(member.id, out var records))
                    continue;

                var intro = records.IntroToSection.results
                    .Where(r => r.status == "awarded" && Key(r.section) == section)
                    .OrderByDescending(r => r.status_updated)
                    .FirstOrDefault();
                if (intro == null)
                    continue;

                model.InvestedYouth++;
                var invested = ScoutingHistoryBuilder.AwardedDate(intro.imported.date_awarded, intro.status_updated);
                if (invested.Date < input.From.Date || invested.Date > input.To.Date)
                    continue;

                model.InvestedInPeriod++;
                model.NewMembers.Add(new NewMember { Name = Name(member), Origin = OriginOf(input, records, rank) });
            }
            model.NewMembers = model.NewMembers.OrderBy(n => n.Name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        // Where a new member came from, from what Terrain holds for them: records submitted under another unit of this
        // group mean they linked up (or moved up) here, records only under units outside it mean they joined from another
        // group, and nothing at all means they are new to Scouting.
        private static NewMemberOrigin OriginOf(SectionReportInput input, SectionMemberRecords records, int rank)
        {
            var elsewhere = records.AllAchievements.results
                .Select(r => r.latest_submission?.unit_id ?? "")
                .Where(unit => unit != "" && unit != input.UnitId)
                .Distinct()
                .ToList();
            var earlierSection = records.AllAchievements.results.Any(r =>
            {
                var recordRank = Array.IndexOf(SectionOrder, Key(r.section));
                return recordRank >= 0 && recordRank < rank;
            });

            if (elsewhere.Any(unit => input.GroupUnitIds.Contains(unit)))
                return NewMemberOrigin.YoungerSection;
            if (input.GroupUnitIds.Count > 0 && elsewhere.Count > 0)
                return NewMemberOrigin.AnotherGroup;
            // records with no unit on them (imported ones) or no way to tell the group's units apart: assume this group
            return earlierSection ? NewMemberOrigin.YoungerSection : NewMemberOrigin.OffTheStreet;
        }

        // The items given to new members: everyone new to Scouting gets the full set, and someone from another group the
        // first four. Counted per item.
        private static List<UnitBadgeCount> InvestitureItems(SectionReportModel model)
        {
            var counts = new Dictionary<string, int>();
            foreach (var item in OffTheStreetItems)
                counts[item] = model.NewOffTheStreet;
            foreach (var item in AnotherGroupItems)
                counts[item] = counts.GetValueOrDefault(item) + model.NewFromAnotherGroup;
            return counts.Where(c => c.Value > 0).Select(c => new UnitBadgeCount { Name = c.Key, Count = c.Value }).ToList();
        }

        // Who is no longer in the unit: people now in the next section's unit, anyone on the roster saved by the last report
        // who has gone, and anyone who came to the unit's events in the period but is not a member now.
        private static void AddDepartures(SectionReportInput input, SectionReportModel model, HashSet<string> currentKeys, string section)
        {
            var departed = new Dictionary<string, DepartedMember>(StringComparer.OrdinalIgnoreCase);
            var nextSection = model.NextSection;
            var nextKeys = input.NextSectionMembers.Select(n => PersonKey(n.Person.MemberNumber, n.Person.Id)).ToHashSet(StringComparer.OrdinalIgnoreCase);

            DepartedMember Add(string id, string number, string name, string source)
            {
                var key = PersonKey(number, id);
                if (!departed.TryGetValue(key, out var entry))
                {
                    entry = new DepartedMember { Id = id, MemberNumber = number, Name = name, Source = source };
                    // someone now in the next section's unit has moved up
                    if (nextKeys.Contains(key))
                    {
                        entry.Status = DepartureStatus.MovedUp;
                        entry.Source = $"Now in {ScoutingHistoryBuilder.Title(nextSection)}s";
                    }
                    departed[key] = entry;
                }
                return entry;
            }

            // Moving up in the period: invested in the next section in the period, with records from this section
            foreach (var next in input.NextSectionMembers)
            {
                var key = PersonKey(next.Person.MemberNumber, next.Person.Id);
                if (currentKeys.Contains(key))
                    continue;
                var intro = next.Records.IntroToSection.results
                    .Where(r => r.status == "awarded" && Key(r.section) == nextSection)
                    .OrderByDescending(r => r.status_updated)
                    .FirstOrDefault();
                if (intro == null)
                    continue;
                var invested = ScoutingHistoryBuilder.AwardedDate(intro.imported.date_awarded, intro.status_updated);
                if (invested.Date < input.From.Date || invested.Date > input.To.Date)
                    continue;
                if (next.Records.AllAchievements.results.Any(r => Key(r.section) == section))
                    Add(next.Person.Id, next.Person.MemberNumber, next.Person.Name, "");
            }

            if (input.Previous != null)
            {
                var when = input.Previous.SavedOn.ToString("d MMM yyyy", System.Globalization.CultureInfo.InvariantCulture);
                foreach (var person in input.Previous.Youth.Where(p => !currentKeys.Contains(PersonKey(p.MemberNumber, p.Id))))
                    Add(person.Id, person.MemberNumber, person.Name, $"On the roster saved {when}");
            }

            foreach (var evt in EventsInPeriod(input).Where(e => e.Scope != "group"))
            {
                var leaderKeys = (evt.Event.attendance?.leader_members ?? new Attendance_Members[0])
                    .Select(a => PersonKey(a.member_number, a.id)).ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (var person in Attendees(evt.Event))
                {
                    var key = PersonKey(person.member_number, person.id);
                    if (currentKeys.Contains(key) || leaderKeys.Contains(key))
                        continue;
                    Add(person.id, person.member_number, $"{person.first_name} {person.last_name}".Trim(), "Came to an event in the period");
                }
            }

            model.Departed = departed.Values.OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        // New and departed leaders can only be named against a roster saved by an earlier report, plus leaders who came
        // to events in the period but aren't in the unit now.
        private static void AddLeaderChanges(SectionReportInput input, SectionReportModel model, List<MemberListModel> leaders, HashSet<string> currentKeys)
        {
            var departed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (input.Previous != null)
            {
                var previousKeys = input.Previous.Leaders.Select(p => PersonKey(p.MemberNumber, p.Id)).ToHashSet(StringComparer.OrdinalIgnoreCase);
                model.NewLeaders = leaders
                    .Where(l => !previousKeys.Contains(PersonKey(l.member_number, l.id)))
                    .Select(Name).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var person in input.Previous.Leaders.Where(p => !currentKeys.Contains(PersonKey(p.MemberNumber, p.Id))))
                    departed[PersonKey(person.MemberNumber, person.Id)] = person.Name;
            }

            foreach (var evt in EventsInPeriod(input).Where(e => e.Scope != "group"))
            {
                foreach (var person in evt.Event.attendance?.leader_members ?? new Attendance_Members[0])
                {
                    var key = PersonKey(person.member_number, person.id);
                    if (!currentKeys.Contains(key))
                        departed.TryAdd(key, $"{person.first_name} {person.last_name}".Trim());
                }
            }

            model.DepartedLeaders = departed.Values.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        // Past events (from the start of the period up to now or the end of the period) and the ones still to come
        private static void AddEvents(SectionReportInput input, SectionReportModel model)
        {
            var current = Distinct(input.CurrentMembers).ToDictionary(m => PersonKey(m.member_number, m.id), m => m, StringComparer.OrdinalIgnoreCase);
            ReportEvent ToReportEvent(UnitEventInput item)
            {
                var evt = item.Event;
                var youth = 0;
                var leaders = 0;
                var leaderKeys = (evt.attendance?.leader_members ?? new Attendance_Members[0])
                    .Select(a => PersonKey(a.member_number, a.id)).ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (var person in Attendees(evt))
                {
                    var key = PersonKey(person.member_number, person.id);
                    // a group event also lists other units' members, so only this unit's own are counted
                    if (item.Scope == "group" && !current.ContainsKey(key))
                        continue;
                    var isLeader = current.TryGetValue(key, out var member) ? UnitReportBuilder.IsLeader(member) : leaderKeys.Contains(key);
                    if (isLeader)
                        leaders++;
                    else
                        youth++;
                }
                return new ReportEvent
                {
                    Title = evt.title,
                    Start = evt.start_datetime,
                    ChallengeArea = ScoutingHistoryBuilder.Title(evt.challenge_area),
                    YouthAttended = youth,
                    LeadersAttended = leaders
                };
            }

            model.PastEvents = EventsInPeriod(input).Select(ToReportEvent).OrderBy(e => e.Start).ToList();
            model.UpcomingEvents = DistinctEvents(input)
                .Where(e => e.Event.start_datetime > input.AsOf && e.Event.start_datetime.Date >= input.From.Date && e.Event.start_datetime.Date <= input.To.Date)
                .Select(ToReportEvent).OrderBy(e => e.Start).ToList();
        }

        // Badges awarded in the period (an approved one has not been handed over yet) to the unit's current youth while in
        // this section
        private static void AddBadges(SectionReportInput input, SectionReportModel model, List<MemberListModel> youth, string section)
        {
            bool InSection(string recordSection) => Key(recordSection) == section;
            bool InPeriod(string imported, DateTime updated)
            {
                var date = ScoutingHistoryBuilder.AwardedDate(imported, updated);
                return date >= input.From.Date && date <= input.To.Date;
            }

            var youthIds = youth.Select(m => m.id).ToHashSet();
            var milestones = new SortedDictionary<int, int>();
            var oas = new SortedDictionary<(string stream, int stage), int>();
            var oasProgressions = 0;
            var sia = new SortedDictionary<string, int>();
            foreach (var records in input.Records.Where(r => youthIds.Contains(r.Key)).Select(r => r.Value))
            {
                var badges = records.Badges;
                foreach (var r in badges.Milestones.results.Where(r => r.status == "awarded" && InSection(r.section) && InPeriod(r.imported.date_awarded, r.status_updated)))
                    milestones[r.achievement_meta.stage] = milestones.GetValueOrDefault(r.achievement_meta.stage) + 1;

                var oasAwarded = badges.Oas.results.Where(r => r.status == "awarded" && InSection(r.section) && InPeriod(r.imported.date_awarded, r.status_updated)).ToList();
                oasProgressions += oasAwarded.Sum(r => r.achievement_meta.stage >= 7 ? 2 : 1);

                // Only the highest level awarded to a member in a stream on one day needs a badge (they can jump several at once)
                var oasBadges = oasAwarded
                    .GroupBy(r => (stream: r.achievement_meta.stream, day: ScoutingHistoryBuilder.AwardedDate(r.imported.date_awarded, r.status_updated).Date))
                    .Select(g => (stream: ScoutingHistoryBuilder.Title(g.Key.stream), stage: g.Max(r => r.achievement_meta.stage)));
                foreach (var badge in oasBadges)
                    oas[badge] = oas.GetValueOrDefault(badge) + 1;

                foreach (var r in badges.Sia.results.Where(r => r.status == "awarded" && InSection(r.section) && InPeriod(r.imported.date_awarded, r.status_updated)))
                {
                    var area = SiaName((r.answers.special_interest_area_selection ?? "").Replace("sia_", ""));
                    sia[area] = sia.GetValueOrDefault(area) + 1;
                }

                model.Badges.PeakAwards += badges.Peak.results.Count(r => r.status == "awarded" && InSection(r.section) && InPeriod("", r.status_updated));
            }

            var awards = new SortedDictionary<string, int>();
            foreach (var r in input.AdditionalAwards.results.Where(r => r.status == "awarded" && youthIds.Contains(r.member_id) && InSection(r.section) && InPeriod(r.imported.date_awarded, r.status_updated)))
            {
                var name = input.AwardNames.TryGetValue(r.achievement_meta.additional_award_id, out var title)
                    ? title
                    : ScoutingHistoryBuilder.Title(r.achievement_meta.additional_award_id);
                awards[name] = awards.GetValueOrDefault(name) + 1;
            }

            model.Badges.OasProgressions = oasProgressions;
            model.Badges.InvestitureItems = InvestitureItems(model);
            model.Badges.Milestones = milestones.Select(m => new UnitBadgeCount { Name = $"Milestone {m.Key}", Count = m.Value }).ToList();
            model.Badges.OasStages = oas.Select(o => new OasStageCount { Stream = o.Key.stream, Stage = o.Key.stage, Count = o.Value }).ToList();
            model.Badges.SiaCategories = sia.Select(s => new UnitBadgeCount { Name = s.Key, Count = s.Value }).ToList();
            model.Badges.AdditionalAwards = awards.Select(a => new UnitBadgeCount { Name = a.Key, Count = a.Value }).ToList();
        }

        // Events once each (the same event can come back from two calendars), not cancelled
        private static IEnumerable<UnitEventInput> DistinctEvents(SectionReportInput input) =>
            input.Events
                .GroupBy(e => e.Event.id)
                .Select(g => g.First())
                .Where(e => !e.Event.status.Contains("cancel", StringComparison.OrdinalIgnoreCase));

        private static IEnumerable<UnitEventInput> EventsInPeriod(SectionReportInput input) =>
            DistinctEvents(input).Where(e =>
                e.Event.start_datetime.Date >= input.From.Date
                && e.Event.start_datetime.Date <= input.To.Date
                && e.Event.start_datetime <= input.AsOf);

        // Everyone recorded as attending, once
        private static IEnumerable<Attendance_Members> Attendees(GetEventResultModel evt)
        {
            var attendance = evt.attendance;
            return (attendance?.attendee_members ?? new Attendance_Members[0])
                .Concat(attendance?.participant_members ?? new Attendance_Members[0])
                .Concat(attendance?.assistant_members ?? new Attendance_Members[0])
                .Concat(attendance?.leader_members ?? new Attendance_Members[0])
                .GroupBy(a => PersonKey(a.member_number, a.id), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First());
        }

        private static RosterPerson Person(MemberListModel member) =>
            new RosterPerson { Id = member.id, MemberNumber = member.member_number, Name = Name(member) };

        private static string Name(MemberListModel member) => $"{member.first_name} {member.last_name}".Trim();
        private static string Key(string? section) => (section ?? "").Trim().ToLowerInvariant();
        // The unit's leaders. Terrain gives the adult leader duty to others in the unit too: youth helpers, and people whose
        // role is in the group (a group leader or advisor). The unit's own leaders are the adults in its "Adult Leaders"
        // patrol; when nobody has a patrol that can't be told, so every adult counts.
        private static List<MemberListModel> LeadersOf(IEnumerable<MemberListModel> members)
        {
            var adults = members.Where(m => UnitReportBuilder.IsLeader(m) && !IsUnder18(m)).ToList();
            var inAPatrol = adults.Where(m => !string.IsNullOrEmpty(m.patrol_name)).ToList();
            return inAPatrol.Count > 0 ? inAPatrol : adults;
        }

        // The areas as Scouts names them
        private static string SiaName(string key) => key switch
        {
            "adventure_sport" => "Adventure & Sport",
            "art_literature" => "Art & Literature",
            "better_world" => "Creating a Better World",
            "environment" => "Environment",
            "growth_development" => "Growth & Development",
            "stem_innovation" => "STEM & Innovation",
            "" => "Other",
            _ => ScoutingHistoryBuilder.Title(key)
        };

        // The ages the sections cover, in whole years
        private static readonly Dictionary<string, (int from, int to)> SectionAges = new Dictionary<string, (int, int)>
        {
            ["joey"] = (5, 7), ["cub"] = (8, 10), ["scout"] = (11, 14), ["venturer"] = (15, 17), ["rover"] = (18, 25)
        };

        // "13y 4m" as years
        private static double? AgeInYears(MemberListModel member)
        {
            var match = System.Text.RegularExpressions.Regex.Match(member.age ?? "", @"^(\d+)y(?:\s*(\d+)m)?");
            if (!match.Success)
                return null;
            return int.Parse(match.Groups[1].Value) + (match.Groups[2].Success ? int.Parse(match.Groups[2].Value) / 12.0 : 0);
        }

        public static AgeProfile AgeProfileOf(IEnumerable<MemberListModel> youth, string section)
        {
            var ages = youth.Select(AgeInYears).Where(a => a != null).Select(a => a!.Value).OrderBy(a => a).ToList();
            var profile = new AgeProfile { Counted = ages.Count };
            if (ages.Count == 0)
                return profile;

            var (from, to) = SectionAges.TryGetValue(section, out var range) ? range : ((int)ages.First(), (int)ages.Last());
            from = Math.Min(from, (int)ages.First());
            to = Math.Max(to, (int)ages.Last());
            profile.AverageYears = ages.Average();
            profile.YoungestYears = ages.First();
            profile.OldestYears = ages.Last();
            profile.Buckets = Enumerable.Range(from, to - from + 1)
                .Select(age => new AgeBucket { Age = age, Count = ages.Count(a => (int)a == age) }).ToList();

            if (ages.Count >= 3 && profile.Buckets.Count > 1)
            {
                // how close the spread is to the same number at every age (entropy against the most it could be)
                var entropy = -profile.Buckets.Where(k => k.Count > 0).Sum(k => (double)k.Count / ages.Count * Math.Log((double)k.Count / ages.Count));
                profile.Evenness = entropy / Math.Log(profile.Buckets.Count);
                var middle = (from + to + 1) / 2.0;
                profile.Lean = (profile.AverageYears - middle) / ((to - from + 1) / 2.0);
            }
            return profile;
        }

        private static bool IsUnder18(MemberListModel member) => int.TryParse(member.age.Split('y')[0], out var years) && years < 18;

        private static string PersonKey(string memberNumber, string id) => string.IsNullOrEmpty(memberNumber) ? id : memberNumber;
    }
}
