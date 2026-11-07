using Topo.Model.History;
using Topo.Model.Members;
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

            // What Terrain holds for a member is read once, even when they matter to two units' reports
            var records = new Dictionary<string, SectionMemberRecords>();
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
                for (var i = 0; i < youth.Count; i++)
                {
                    progress?.Invoke($"{unitName}: members' records {i + 1} of {youth.Count}");
                    input.Records[youth[i].id] = await GetRecords(youth[i].id, records, withBadges: true);
                }

                progress?.Invoke($"{unitName}: additional awards");
                input.AdditionalAwards = await _terrainAPIService.GetUnitAdditionalAwardAchievements(unitId);

                // The next section's unit, if the person can see it, shows who has moved up
                var nextSection = SectionReportBuilder.NextSectionOf(section);
                foreach (var nextUnit in groupUnits.Where(u => nextSection != "" && string.Equals(u.section, nextSection, StringComparison.OrdinalIgnoreCase)))
                {
                    progress?.Invoke($"{unitName}: {nextUnit.name}");
                    var nextMembers = (await _membersService.GetMembersAsync(nextUnit.id)).Where(m => !UnitReportBuilder.IsLeader(m)).ToList();
                    foreach (var next in nextMembers)
                    {
                        input.NextSectionMembers.Add(new NextSectionMember
                        {
                            Person = new RosterPerson { Id = next.id, MemberNumber = next.member_number, Name = $"{next.first_name} {next.last_name}".Trim() },
                            Records = await GetRecords(next.id, records, withBadges: false)
                        });
                    }
                }

                result.Report.Sections.Add(SectionReportBuilder.Build(input));
                result.Rosters[unitId] = SectionReportBuilder.TakeSnapshot(members, now);
            }

            return result;
        }

        // The introduction and every record's section and unit for a member, plus their badges when asked for
        private async Task<SectionMemberRecords> GetRecords(string memberId, Dictionary<string, SectionMemberRecords> cache, bool withBadges)
        {
            if (!cache.TryGetValue(memberId, out var memberRecords))
            {
                memberRecords = new SectionMemberRecords
                {
                    IntroToSection = await _terrainAPIService.GetIntroductionToSectionResultsForMember(memberId),
                    AllAchievements = await _terrainAPIService.GetAllAchievementsForMember(memberId)
                };
                cache[memberId] = memberRecords;
            }

            if (withBadges && !memberRecords.BadgesLoaded)
            {
                memberRecords.BadgesLoaded = true;
                memberRecords.Badges = new UnitMemberBadges
                {
                    Milestones = await _terrainAPIService.GetMilestoneResultsForMember(memberId),
                    Oas = await _terrainAPIService.GetOASResultsForMember(memberId),
                    Sia = await _terrainAPIService.GetSIAResultsForMember(memberId),
                    Peak = await _terrainAPIService.GetPeakAwardResultsForMember(memberId)
                };
            }
            return memberRecords;
        }
    }
}
