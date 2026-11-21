using System.Globalization;
using Topo.Model;
using Topo.Model.Logbook;

namespace Topo.Services
{
    public interface ILogbookService
    {
        public Task<List<MemberLogbookReportViewModel>> GenerateLogbookData(List<KeyValuePair<string, string>> selectedMembers);
    }
    public class LogbookService : ILogbookService
    {
        private readonly ITerrainAPIService _terrainAPIService;

        public LogbookService(ITerrainAPIService terrainAPIService)
        {
            _terrainAPIService = terrainAPIService;
        }

        public async Task<List<MemberLogbookReportViewModel>> GenerateLogbookData(List<KeyValuePair<string, string>> selectedMembers)
        {
            var memberLogbooks = new List<MemberLogbookReportViewModel>();
            foreach (var memberKVP in selectedMembers)
            {
                await _terrainAPIService.RevokeAssumedProfiles();
                await _terrainAPIService.AssumeProfile(memberKVP.Key);
                var getMemberLogbookMetrics = await _terrainAPIService.GetMemberLogbookMetrics(memberKVP.Key);
                var totalNightsCamped = getMemberLogbookMetrics.results.Where(r => r.name == "total_nights_camped").FirstOrDefault()?.value ?? 0;
                var totalKmsHiked = (getMemberLogbookMetrics.results.Where(r => r.name == "total_distance_hiked").FirstOrDefault()?.value ?? 0) / 1000.0f;
                var getMemberLogbookSummary = await _terrainAPIService.GetMemberLogbookSummary(memberKVP.Key);
                if (getMemberLogbookSummary != null)
                {
                    // The summary only lists the entries; each one's detail is a request of its own. They are read a few at a time
                    // (all while this Scout's profile is assumed) and come back in the order listed, so the report is unchanged.
                    var logbookDetails = await ReadsAtOnce.SelectAsync(getMemberLogbookSummary.results,
                        summaryResult => _terrainAPIService.GetMemberLogbookDetail(memberKVP.Key, summaryResult.id));
                    foreach (var getLogbookDetail in logbookDetails)
                    {
                        if (getLogbookDetail != null)
                        {
                            var activityDate = DateTime.ParseExact(getLogbookDetail.start_date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            var endDate = DateTime.ParseExact(getLogbookDetail.end_date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            var nightsCamped = getLogbookDetail.categories.Any(c => c == "camping") ?
                                (endDate - activityDate).TotalDays :
                                0;
                            var kmsHiked = getLogbookDetail.categories.Any(c => c == "walking_hike") ?
                                getLogbookDetail.distance_walkabout / 1000.0f :
                                0;
                            memberLogbooks.Add(new MemberLogbookReportViewModel
                            {
                                MemberName = memberKVP.Value,
                                TotalKilometersHiked = totalKmsHiked,
                                TotalNightsCamped = totalNightsCamped,
                                ActivityName = getLogbookDetail.title,
                                ActivityArea = getLogbookDetail.achievement_meta.branch,
                                ActivityDate = activityDate,
                                ActivityLead = getLogbookDetail.details.who_lead,
                                MemberRole = getLogbookDetail.details.your_role,
                                KilometersHiked = kmsHiked,
                                NightsCamped = (float)nightsCamped,
                                VerifierName = getLogbookDetail.details.verifier.name,
                                VerifierContact = getLogbookDetail.details.verifier.contact,
                                Weather = getLogbookDetail.details.weather,
                                Location = getLogbookDetail.details.location,
                                Grade = getLogbookDetail.details.activity_grade,
                                OtherParticipants = getLogbookDetail.details.other_participants,
                                Purpose = getLogbookDetail.details.purpose
                            });
                        }
                    }
                }
            }
            await _terrainAPIService.RevokeAssumedProfiles();

            return memberLogbooks;
        }
    }
}



