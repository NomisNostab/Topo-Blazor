using Topo.Model.UnitReport;

namespace Topo.Services
{
    public interface IUnitReportService
    {
        public Task<UnitReportModel> GetUnitReport(string unitId, DateTime from, DateTime to, string[] calendarIds, Action<string>? progress = null);
    }

    public class UnitReportService : IUnitReportService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly IMembersService _membersService;
        private readonly IProgramService _programService;

        // award id -> title; the same for everyone so fetched once
        private static Dictionary<string, string>? _awardNames;

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
            input.Events = await GetEvents(from, to, calendarIds);

            // Badges come from each current youth member's records
            var youth = input.CurrentMembers.Where(m => !UnitReportBuilder.IsLeader(m)).ToList();
            for (var i = 0; i < youth.Count; i++)
            {
                progress?.Invoke($"badges {i + 1} of {youth.Count}");
                input.BadgesByMember[youth[i].id] = new UnitMemberBadges
                {
                    Milestones = await _terrainAPIService.GetMilestoneResultsForMember(youth[i].id),
                    Oas = await _terrainAPIService.GetOASResultsForMember(youth[i].id),
                    Sia = await _terrainAPIService.GetSIAResultsForMember(youth[i].id),
                    Peak = await _terrainAPIService.GetPeakAwardResultsForMember(youth[i].id)
                };
            }

            progress?.Invoke("additional awards");
            input.AdditionalAwards = await _terrainAPIService.GetUnitAdditionalAwardAchievements(unitId);
            input.AwardNames = await GetAwardNames();

            return UnitReportBuilder.Build(input);
        }

        // Terrain only shows the calendars the user has switched on, so they are switched to the chosen ones and back
        private async Task<List<UnitEventInput>> GetEvents(DateTime from, DateTime to, string[] calendarIds)
        {
            await _programService.SetCalendar(calendarIds);
            try
            {
                return await _programService.GetEventDetailsForDates(from, to);
            }
            finally
            {
                await _programService.ResetCalendar();
            }
        }

        private async Task<Dictionary<string, string>> GetAwardNames()
        {
            if (_awardNames == null)
            {
                var specifications = await _terrainAPIService.GetAdditionalAwardSpecifications();
                _awardNames = specifications.AwardDescriptions.GroupBy(a => a.id).ToDictionary(g => g.Key, g => g.First().title);
            }
            return _awardNames;
        }
    }
}
