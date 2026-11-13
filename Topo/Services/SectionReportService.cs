using Topo.Model.History;
using Topo.Model.Milestone;
using Topo.Model.Members;
using Topo.Model.Progress;
using Topo.Model.SectionReport;
using Topo.Model.UnitReport;

namespace Topo.Services
{
    public class SectionReportRequest
    {
        public List<string> UnitIds { get; set; } = new List<string>();
        public bool IsGroup { get; set; }
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        // The calendars to read each unit's events from
        public Dictionary<string, string[]> CalendarIdsByUnit { get; set; } = new Dictionary<string, string[]>();
        // The roster saved by each unit's last report, if there is one
        public Dictionary<string, RosterSnapshot> PreviousRosters { get; set; } = new Dictionary<string, RosterSnapshot>();
    }

    public class SectionReportResult
    {
        public SectionReportSet Report { get; set; } = new SectionReportSet();
        // Each unit's roster as it is now, to save for the next report
        public Dictionary<string, RosterSnapshot> Rosters { get; set; } = new Dictionary<string, RosterSnapshot>();
    }

    public interface ISectionReportService
    {
        public Task<SectionReportResult> GetSectionReport(SectionReportRequest request, Action<string>? progress = null);
    }

    public class SectionReportService : ISectionReportService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly IMembersService _membersService;
        private readonly IUnitReportService _unitReportService;

        public SectionReportService(StorageService storageService, ITerrainAPIService terrainAPIService, IMembersService membersService, IUnitReportService unitReportService)
        {
            _storageService = storageService;
            _terrainAPIService = terrainAPIService;
            _membersService = membersService;
            _unitReportService = unitReportService;
        }

        public async Task<SectionReportResult> GetSectionReport(SectionReportRequest request, Action<string>? progress = null)
        {
            var now = DateTime.Now;
            var result = new SectionReportResult
            {
                Report = new SectionReportSet
                {
                    GeneratedAt = now,
                    GeneratedBy = _storageService.MemberName ?? "",
                    GroupName = _storageService.GroupName ?? "",
                    From = request.From,
                    To = request.To,
                    IsGroup = request.IsGroup
                }
            };

            // Every unit of the group, so a unit's next section (where its members move up to) can be found
            var groupUnits = string.IsNullOrEmpty(_storageService.GroupId)
                ? new List<GroupUnit>()
                : await _terrainAPIService.GetGroupUnitsAsync(_storageService.GroupId);
            var awardNames = await AdditionalAwardNames.Get(_terrainAPIService);

            // What Terrain holds for a member is read once, even when they matter to two units' reports. Each kind of record is
            // read once for a whole unit (every record carries its member_id) instead of once per member, then handed out by member.
            var records = new Dictionary<string, SectionMemberRecords>();
            var introsByUnit = new Dictionary<string, ILookup<string, Topo.Model.Progress.Result>>();
            var badgesByUnit = new Dictionary<string, UnitBadgeRecords>();
            var eventsByCalendars = new Dictionary<string, List<UnitEventInput>>();

            foreach (var unitId in request.UnitIds)
            {
                var profile = _storageService.GetProfilesResult?.profiles?.FirstOrDefault(p => p.unit?.id == unitId);
                var unitName = profile?.unit?.name ?? "";
                var section = profile?.unit?.section ?? "";
                progress?.Invoke($"{unitName}: members");
                var members = await _membersService.GetMembersAsync(unitId);

                var input = new SectionReportInput
                {
                    UnitId = unitId,
                    UnitName = unitName,
                    Section = section,
                    From = request.From,
                    To = request.To,
                    AsOf = now,
                    CurrentMembers = members,
                    AwardNames = awardNames,
                    GroupUnitIds = groupUnits.Select(u => u.id).ToHashSet(),
                    Previous = request.PreviousRosters.GetValueOrDefault(unitId)
                };

                progress?.Invoke($"{unitName}: events");
                var calendarIds = request.CalendarIdsByUnit.GetValueOrDefault(unitId) ?? new string[0];
                var calendarKey = string.Join(",", calendarIds.OrderBy(c => c));
                if (calendarIds.Length > 0 && !eventsByCalendars.ContainsKey(calendarKey))
                {
                    eventsByCalendars[calendarKey] = await _unitReportService.GetEventsForCalendars(request.From, request.To, calendarIds);
                }
                if (calendarIds.Length > 0)
                    input.Events = eventsByCalendars[calendarKey];

                var youth = members.Where(m => !UnitReportBuilder.IsLeader(m)).ToList();
                // A unit with no youth (an adults-only team) has no records to read
                if (youth.Count > 0)
                {
                    progress?.Invoke($"{unitName}: members' records");
                    var intros = await IntrosOf(unitId, introsByUnit);
                    var badges = await BadgesOf(unitId, badgesByUnit);
                    foreach (var member in youth)
                        input.Records[member.id] = await GetRecords(member.id, intros, badges, section, request, records);
                }

                progress?.Invoke($"{unitName}: additional awards");
                input.AdditionalAwards = await _terrainAPIService.GetUnitAdditionalAwardAchievements(unitId);

                // The next section's unit, if the person can see it, shows who has moved up
                var nextSection = SectionReportBuilder.NextSectionOf(section);
                foreach (var nextUnit in groupUnits.Where(u => nextSection != "" && string.Equals(u.section, nextSection, StringComparison.OrdinalIgnoreCase)))
                {
                    progress?.Invoke($"{unitName}: {nextUnit.name}");
                    var nextMembers = (await _membersService.GetMembersAsync(nextUnit.id)).Where(m => !UnitReportBuilder.IsLeader(m)).ToList();
                    var nextIntros = nextMembers.Count > 0 ? await IntrosOf(nextUnit.id, introsByUnit) : null;
                    foreach (var next in nextMembers)
                    {
                        input.NextSectionMembers.Add(new NextSectionMember
                        {
                            Person = new RosterPerson { Id = next.id, MemberNumber = next.member_number, Name = $"{next.first_name} {next.last_name}".Trim() },
                            Records = await GetRecords(next.id, nextIntros!, null, nextSection, request, records)
                        });
                    }
                }

                result.Report.Sections.Add(SectionReportBuilder.Build(input));
                result.Rosters[unitId] = SectionReportBuilder.TakeSnapshot(members, now);
            }

            return result;
        }

        // A unit's Introductions to Section and its youth's badge records, each read once for the whole unit
        private async Task<ILookup<string, Topo.Model.Progress.Result>> IntrosOf(string unitId, Dictionary<string, ILookup<string, Topo.Model.Progress.Result>> cache)
        {
            if (!cache.TryGetValue(unitId, out var intros))
            {
                intros = (await _terrainAPIService.GetUnitAchievementsOfType<GetIntroductionResultsModel>(unitId, "intro_section")).results.ToLookup(r => r.member_id);
                cache[unitId] = intros;
            }
            return intros;
        }

        private async Task<UnitBadgeRecords> BadgesOf(string unitId, Dictionary<string, UnitBadgeRecords> cache)
        {
            if (!cache.TryGetValue(unitId, out var badges))
            {
                badges = new UnitBadgeRecords
                {
                    Milestones = (await _terrainAPIService.GetUnitAchievementsOfType<GetMilestoneResultModel>(unitId, "milestone")).results.ToLookup(r => r.member_id),
                    Oas = (await _terrainAPIService.GetUnitAchievementsOfType<Topo.Model.OAS.GetUnitAchievementsResultsModel>(unitId, "outdoor_adventure_skill")).results.ToLookup(r => r.member_id),
                    Sia = (await _terrainAPIService.GetUnitAchievementsOfType<Topo.Model.SIA.GetSIAResultsModel>(unitId, "special_interest_area")).results.ToLookup(r => r.member_id),
                    Peak = (await _terrainAPIService.GetUnitAchievementsOfType<GetPeakAwardsResultsModel>(unitId, "peak_award")).results.ToLookup(r => r.member_id)
                };
                cache[unitId] = badges;
            }
            return badges;
        }

        private class UnitBadgeRecords
        {
            public ILookup<string, MilestoneResult> Milestones { get; set; } = null!;
            public ILookup<string, Topo.Model.OAS.Result> Oas { get; set; } = null!;
            public ILookup<string, Topo.Model.SIA.Result> Sia { get; set; } = null!;
            public ILookup<string, PeakAwardResult> Peak { get; set; } = null!;
        }

        // A member's introduction and badges from the unit-wide reads. Terrain's untyped "everything for this member" read is
        // the one that stays per member, because it is only needed for people who are new in the period: the report uses it
        // to say where a new member came from.
        private async Task<SectionMemberRecords> GetRecords(string memberId, ILookup<string, Topo.Model.Progress.Result> intros, UnitBadgeRecords? badges, string section,
            SectionReportRequest request, Dictionary<string, SectionMemberRecords> cache)
        {
            if (!cache.TryGetValue(memberId, out var memberRecords))
            {
                memberRecords = new SectionMemberRecords
                {
                    IntroToSection = new GetIntroductionResultsModel { results = intros[memberId].ToArray() }
                };
                cache[memberId] = memberRecords;
            }

            if (!memberRecords.AllAchievementsLoaded && SectionReportBuilder.InvestedInPeriod(memberRecords.IntroToSection, section, request.From, request.To))
            {
                memberRecords.AllAchievementsLoaded = true;
                memberRecords.AllAchievements = await _terrainAPIService.GetAllAchievementsForMember(memberId);
            }

            if (badges != null && !memberRecords.BadgesLoaded)
            {
                memberRecords.BadgesLoaded = true;
                memberRecords.Badges = new UnitMemberBadges
                {
                    Milestones = new GetMilestoneResultModel { results = badges.Milestones[memberId].ToArray() },
                    Oas = new Topo.Model.OAS.GetUnitAchievementsResultsModel { results = badges.Oas[memberId].ToArray() },
                    Sia = new Topo.Model.SIA.GetSIAResultsModel { results = badges.Sia[memberId].ToArray() },
                    Peak = new GetPeakAwardsResultsModel { results = badges.Peak[memberId].ToArray() }
                };
            }
            return memberRecords;
        }
    }
}
