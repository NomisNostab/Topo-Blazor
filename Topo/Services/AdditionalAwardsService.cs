﻿using System.Globalization;
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
                    })
                    .ToList();
                _storageService.AdditionalAwardSpecifications = awardSpecificationsList;
            }
            var unitAchievementsResult = await _terrainAPIService.GetUnitAdditionalAwardAchievements(selectedUnitId ?? "");
            var approvedAwards = new List<ApprovalsListModel>(); // await _approvalsService.GetApprovalListItems(selectedUnitId ?? "");
            var additionalAwardsList = new List<AdditionalAwardListModel>();
            var lastMemberProcessed = "";
            var memberName = "";
            foreach(var memberKVP in selectedMembers)
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
                    MemberName = memberName,
                    AwardId = "",
                    AwardName = "",
                    AwardSortIndex = 0,
                    AwardDate = null,
                    PresentedDate = null
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
                        MemberName = memberName,
                        AwardId = awardSpecification?.id ?? "",
                        AwardName = awardSpecification?.name ?? "",
                        AwardSortIndex = awardSpecification?.additionalAwardSortIndex ?? 0,
                        AwardDate = awardStatusDate,
                        PresentedDate = awardPresentedDate ?? awardStatusDate //null
                    });
               }
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
    }
}
