using System.Globalization;
using Topo.Model.AdditionalAwards;
using Topo.Model.Approvals;

namespace Topo.Services
{
    public interface IAdditionalAwardService
    {
        public Task<AdditionalAwardsReportDataModel> GenerateAdditionalAwardReportData(string selectedUnitId, List<KeyValuePair<string, string>> selectedMembers);
    }
    public class AdditionalAwardService : IAdditionalAwardService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly IReportService _reportService;
        //private readonly IApprovalsService _approvalsService;

        public AdditionalAwardService(ITerrainAPIService terrainAPIService,
            StorageService storageService,
            IReportService reportService) //, IApprovalsService approvalsService)
        {
            _terrainAPIService = terrainAPIService;
            _storageService = storageService;
            _reportService = reportService;
            //_approvalsService = approvalsService;
        }
        public async Task<AdditionalAwardsReportDataModel> GenerateAdditionalAwardReportData(string selectedUnitId, List<KeyValuePair<string, string>> selectedMembers)
        {
            var awardSpecificationsList = _storageService.AdditionalAwardSpecifications;
            if (awardSpecificationsList == null || awardSpecificationsList.Count == 0)
            {
                var additionalAwardsSpecifications = await _terrainAPIService.GetAdditionalAwardSpecifications();
                var additionalAwardSortIndex = 0;
                awardSpecificationsList = additionalAwardsSpecifications.AwardDescriptions
                    .Select(x => new AdditionalAwardSpecificationListModel()
                    {
                        id = x.id,
                        name = x.title,
                        additionalAwardSortIndex = additionalAwardSortIndex++,
                        additionalAwardValue = GetAwardValue(x.id)
                    })
                    .ToList();
                _storageService.AdditionalAwardSpecifications = awardSpecificationsList;
            }
            var unitAchievementsResult = await _terrainAPIService.GetUnitAdditionalAwardAchievements(selectedUnitId ?? "");
            var approvedAwards = new List<ApprovalsListModel>(); // await _approvalsService.GetApprovalListItems(selectedUnitId ?? "");
            var additionalAwardsList = new List<AdditionalAwardListModel>();
            var lastMemberProcessed = "";
            var memberName = "";
            foreach (var memberKVP in selectedMembers)
            {
                await _terrainAPIService.RevokeAssumedProfiles();
                await _terrainAPIService.AssumeProfile(memberKVP.Key);
                var getMemberLogbookMetrics = await _terrainAPIService.GetMemberLogbookMetrics(memberKVP.Key);
                var totalNightsCamped = getMemberLogbookMetrics.results.Where(r => r.name == "total_nights_camped").FirstOrDefault()?.value ?? 0;
                var totalKmsHiked = (getMemberLogbookMetrics.results.Where(r => r.name == "total_distance_hiked").FirstOrDefault()?.value ?? 0) / 1000.0f;
                memberName = $"{memberKVP.Value} ({totalNightsCamped} Nights, {totalKmsHiked} KMs)";
                lastMemberProcessed = memberKVP.Key;

                additionalAwardsList.Add(new AdditionalAwardListModel
                {
                    MemberName = memberKVP.Value,
                    NightsCamped = (int)totalNightsCamped,
                    KMsHiked = (int)totalKmsHiked,
                    AwardId = "",
                    AwardName = "",
                    AwardSortIndex = 0,
                    AwardDate = null,
                    PresentedDate = null,
                    AwardUnits = 0
                });

                foreach (var result in unitAchievementsResult.results.Where(r => r.member_id == memberKVP.Key))
                {
                    var awardSpecification = awardSpecificationsList.Where(a => a.id == result.achievement_meta.additional_award_id).FirstOrDefault();
                    var awardStatus = result.status;
                    var awardStatusDate = result.status_updated;
                    if (result.imported != null)
                        awardStatusDate = DateTime.ParseExact(result.imported.date_awarded, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    var award = approvedAwards.Where(a => a.achievement_id == result.id && a.submission_type.ToLower() == "review").FirstOrDefault();
                    DateTime? awardPresentedDate = new DateTime?();
                    if (award != null)
                    {
                        if (award.presented_date.HasValue)
                        {
                            awardPresentedDate = award.presented_date.Value.ToLocalTime();
                        }
                    }
                    additionalAwardsList.Add(new AdditionalAwardListModel
                    {
                        MemberName = memberKVP.Value,
                        NightsCamped = (int)totalNightsCamped,
                        KMsHiked = (int)totalKmsHiked,
                        AwardId = awardSpecification?.id ?? "",
                        AwardName = awardSpecification?.name ?? "",
                        AwardSortIndex = awardSpecification?.additionalAwardSortIndex ?? 0,
                        AwardDate = awardStatusDate,
                        PresentedDate = awardPresentedDate ?? awardStatusDate, //null
                        AwardUnits = awardSpecification?.additionalAwardValue ?? 0
                    });
                }

                // Get Next Camper/Walkabout Award
                var highestCamperAward = additionalAwardsList.Where(a => a.MemberName == memberKVP.Value && a.AwardId.StartsWith("camper"))
                    .OrderByDescending(a => a.AwardSortIndex)
                    .FirstOrDefault();
                var nextCamperAward = awardSpecificationsList.Where(a => a.additionalAwardSortIndex == highestCamperAward?.AwardSortIndex + 1)
                    .FirstOrDefault();
                additionalAwardsList.Add(new AdditionalAwardListModel
                {
                    MemberName = memberKVP.Value,
                    NightsCamped = (int)totalNightsCamped,
                    KMsHiked = (int)totalKmsHiked,
                    AwardId = nextCamperAward?.id ?? "",
                    AwardName = nextCamperAward?.name ?? "",
                    AwardSortIndex = nextCamperAward?.additionalAwardSortIndex ?? 0,
                    AwardDate = null,
                    PresentedDate = null,
                    AwardUnits = nextCamperAward?.additionalAwardValue ?? 0
                });

                var highestWalkaboutAward = additionalAwardsList.Where(a => a.MemberName == memberKVP.Value && a.AwardId.StartsWith("walkabout"))
                    .OrderByDescending(a => a.AwardSortIndex)
                    .FirstOrDefault();
                var nextWalkaboutAward = awardSpecificationsList.Where(a => a.additionalAwardSortIndex == highestWalkaboutAward?.AwardSortIndex + 1)
                    .FirstOrDefault();
                additionalAwardsList.Add(new AdditionalAwardListModel
                {
                    MemberName = memberKVP.Value,
                    NightsCamped = (int)totalNightsCamped,
                    KMsHiked = (int)totalKmsHiked,
                    AwardId = nextWalkaboutAward?.id ?? "",
                    AwardName = nextWalkaboutAward?.name ?? "",
                    AwardSortIndex = nextWalkaboutAward?.additionalAwardSortIndex ?? 0,
                    AwardDate = null,
                    PresentedDate = null,
                    AwardUnits = nextWalkaboutAward?.additionalAwardValue ?? 0
                });

            }

            await _terrainAPIService.RevokeAssumedProfiles();

            var sortedAdditionalAwardsList = additionalAwardsList.OrderBy(a => a.MemberName).ThenBy(a => a.AwardSortIndex).ToList();
            var distinctAwards = sortedAdditionalAwardsList.OrderBy(x => x.AwardSortIndex).Select(x => x.AwardId).Distinct().ToList();

            var reportData = new AdditionalAwardsReportDataModel();
            reportData.AwardSpecificationsList = awardSpecificationsList;
            reportData.DistinctAwards = distinctAwards;
            reportData.SortedAdditionalAwardsList = sortedAdditionalAwardsList;

            return reportData;
        }

        private int GetAwardValue(string awardId)
        {
            if (string.IsNullOrWhiteSpace(awardId))
            {
                return 0;
            }
            string[] awardBits = awardId.Split('_');
            if (awardBits[0] == "camper")
            {
                bool isInt = int.TryParse(awardBits[2], out int awardUnit);
                if (isInt)
                {
                    return awardUnit;
                }
            }
            if (awardBits[0] == "walkabout")
            {
                bool isInt = int.TryParse(awardBits[1], out int awardUnit);
                if (isInt)
                {
                    return awardUnit;
                }
            }
            return 0;
        }
    }
}
