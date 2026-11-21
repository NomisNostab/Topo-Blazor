using Topo.Model.Milestone;
using Topo.Model.Progress;
using Topo.Model.UnitReport;

namespace Topo.Services
{
    public interface IUnitReportService
    {
        public Task<UnitReportModel> GetUnitReport(string unitId, DateTime from, DateTime to, string[] calendarIds, Action<string>? progress = null);
        // Every event in the range from the chosen calendars, in full (attendance, logbook data) with its scope
        public Task<List<UnitEventInput>> GetEventsForCalendars(DateTime from, DateTime to, string[] calendarIds);
    }

    public class UnitReportService : IUnitReportService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly IMembersService _membersService;
        private readonly IProgramService _programService;

        public UnitReportService(StorageService storageService, ITerrainAPIService terrainAPIService, IMembersService membersService, IProgramService programService)
        {
            _storageService = storageService;
            _terrainAPIService = terrainAPIService;
            _membersService = membersService;
            _programService = programService;
        }

        public async Task<UnitReportModel> GetUnitReport(string unitId, DateTime from, DateTime to, string[] calendarIds, Action<string>? progress = null)
        {
            var profile = _storageService.GetProfilesResult?.profiles?.FirstOrDefault(p => p.unit?.id == unitId);
            var input = new UnitReportInput
            {
                UnitName = profile?.unit?.name ?? "",
                Section = profile?.unit?.section ?? "",
                From = from,
                To = to,
                AsOf = DateTime.Now
            };

            progress?.Invoke("members");
            input.CurrentMembers = await _membersService.GetMembersAsync(unitId);

            progress?.Invoke("events");
            input.Events = await GetEventsForCalendars(from, to, calendarIds);

            // Badges come from each current youth member's records. Read each kind once for the whole unit instead of once per youth:
            // every record carries its member_id, so ToLookup hands back one member's records (none gives an empty list).
            // The records include earlier sections, which the builder filters out.
            var youth = input.CurrentMembers.Where(m => !UnitReportBuilder.IsLeader(m)).ToList();
            // A unit with no youth (an adults-only team) has no badge records to read
            if (youth.Count > 0)
            {
                progress?.Invoke("badges");
                var milestonesByMember = (await _terrainAPIService.GetUnitAchievementsOfType<GetMilestoneResultModel>(unitId, "milestone")).results.ToLookup(r => r.member_id);
                var oasByMember = (await _terrainAPIService.GetUnitAchievementsOfType<Topo.Model.OAS.GetUnitAchievementsResultsModel>(unitId, "outdoor_adventure_skill")).results.ToLookup(r => r.member_id);
                var siaByMember = (await _terrainAPIService.GetUnitAchievementsOfType<Topo.Model.SIA.GetSIAResultsModel>(unitId, "special_interest_area")).results.ToLookup(r => r.member_id);
                var peakByMember = (await _terrainAPIService.GetUnitAchievementsOfType<GetPeakAwardsResultsModel>(unitId, "peak_award")).results.ToLookup(r => r.member_id);
                foreach (var member in youth)
                {
                    input.BadgesByMember[member.id] = new UnitMemberBadges
                    {
                        Milestones = new GetMilestoneResultModel { results = milestonesByMember[member.id].ToArray() },
                        Oas = new Topo.Model.OAS.GetUnitAchievementsResultsModel { results = oasByMember[member.id].ToArray() },
                        Sia = new Topo.Model.SIA.GetSIAResultsModel { results = siaByMember[member.id].ToArray() },
                        Peak = new GetPeakAwardsResultsModel { results = peakByMember[member.id].ToArray() }
                    };
                }
            }

            progress?.Invoke("additional awards");
            input.AdditionalAwards = await _terrainAPIService.GetUnitAdditionalAwardAchievements(unitId);
            input.AwardNames = await AdditionalAwardNames.Get(_terrainAPIService);

            return UnitReportBuilder.Build(input);
        }

        // Terrain only shows the calendars the user has switched on, so they are switched to the chosen ones and back
        public async Task<List<UnitEventInput>> GetEventsForCalendars(DateTime from, DateTime to, string[] calendarIds)
        {
            await _programService.SetCalendar(calendarIds);
            try
            {
                return await GetEventDetails(from, to);
            }
            finally
            {
                await _programService.ResetCalendar();
            }
        }

        // Every event in the range, in full (attendance, logbook data) with its scope. Asked for in 90 day windows
        // so that "all time" is many small requests rather than one very large one.
        private async Task<List<UnitEventInput>> GetEventDetails(DateTime from, DateTime to)
        {
            var userId = _storageService.GetProfilesResult?.profiles?.FirstOrDefault()?.member?.id ?? "";
            var events = new List<UnitEventInput>();
            var seen = new HashSet<string>();
            for (var windowStart = from.Date; windowStart <= to.Date; windowStart = windowStart.AddDays(90))
            {
                var windowEnd = windowStart.AddDays(89) < to.Date ? windowStart.AddDays(89) : to.Date;
                var result = await _terrainAPIService.GetEventsAsync(userId, windowStart.AddDays(-1), windowEnd.AddDays(1));
                foreach (var summary in result?.results ?? new Topo.Model.Program.Event[0])
                {
                    if (!seen.Add(summary.id))
                        continue;
                    var detail = await _terrainAPIService.GetEventAsync(summary.id);
                    if (detail != null)
                        events.Add(new UnitEventInput { Scope = summary.invitee_type, Event = detail });
                }
            }
            return events;
        }
    }
}
