using System.Globalization;
using Topo.Model.Approvals;
using Topo.Model.Milestone;
using Topo.Model.Progress;
using Topo.Model.Wallchart;
using Topo.Pages;

namespace Topo.Services
{
    public interface IWallchartService
    {
        Task<List<WallchartItemModel>> GetWallchartItems(string selectedUnitId, Action<string>? progress = null);
    }
    public class WallchartService : IWallchartService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly IApprovalsService _approvalService;
        private readonly IMembersService _membersService;
        private List<ApprovalsListModel> approvals = new List<ApprovalsListModel>();

        public WallchartService(StorageService storageService, ITerrainAPIService terrainAPIService, IApprovalsService approvalService, IMembersService membersService)
        {
            _storageService = storageService;
            _terrainAPIService = terrainAPIService;
            _approvalService = approvalService;
            _membersService = membersService;
        }

        public async Task<List<WallchartItemModel>> GetWallchartItems(string selectedUnitId, Action<string>? progress = null)
        {
            var members = await _membersService.GetMembersAsync(selectedUnitId);
            var cachedWallchartItems = _storageService.CachedWallchartItems.Where(cm => cm.Key == selectedUnitId).FirstOrDefault().Value;
            if (cachedWallchartItems != null)
            {
                // Update cache member name
                foreach (var item in cachedWallchartItems)
                {
                    var member = members.Where(m => m.id == item.MemberId).FirstOrDefault();
                    item.MemberName = $"{member?.first_name} {_membersService.GetMemberLastName(selectedUnitId, item.MemberId).Result}";
                }
                return cachedWallchartItems;
            }

            var section = _storageService.Section;
            approvals = await _approvalService.ReadApprovalListFromLocalStorage(selectedUnitId);
            await _terrainAPIService.RevokeAssumedProfiles();

            // Terrain works out each member's Peak Award Progress itself; read it rather than estimating it here
            var terrainPeakProgress = (await _terrainAPIService.GetUnitMetrics(selectedUnitId))
                .GroupBy(m => m.member_id)
                .ToDictionary(g => g.Key, g => g.First().achievement_metrics.peak_award_progress);
            var wallchartItems = new List<WallchartItemModel>();
            var youth = members.Where(m => m.isAdultLeader == 0).OrderBy(m => m.last_name).ThenBy(m => m.first_name).ToList();
            foreach (var member in youth)
            {
                progress?.Invoke($"{member.first_name} {member.last_name} ({youth.IndexOf(member) + 1} of {youth.Count})");
                var wallchartItem = new WallchartItemModel();

                // Get Milestone for member
                var memberMilestones = await _terrainAPIService.GetMilestoneResultsForMember(member.id);
                foreach (var milestoneResult in AchievementSelection.OnePerMilestoneStage(memberMilestones.results.Where(r => r.section == _storageService.Section)))
                {
                    MoveMilestoneAssistsToLeads(milestoneResult);
                    switch (milestoneResult.achievement_meta.stage)
                    {
                        case 1:
                            if (milestoneResult.status == "awarded")
                            {
                                wallchartItem.Milestone1Community = 6;
                                wallchartItem.Milestone1Creative = 6;
                                wallchartItem.Milestone1Outdoors = 6;
                                wallchartItem.Milestone1PersonalGrowth = 6;
                                wallchartItem.Milestone1Assist = 0;
                                wallchartItem.Milestone1Lead = 2;
                                wallchartItem.Milestone1Awarded = milestoneResult.status_updated;
                                wallchartItem.Milestone1Presented = GetPresentedDate(member.id, "Milestone 1");
                            }
                            else
                            {
                                wallchartItem.Milestone1Community = (int)milestoneResult.event_count.participant.community;
                                wallchartItem.Milestone1Creative = (int)milestoneResult.event_count.participant.creative;
                                wallchartItem.Milestone1Outdoors = (int)milestoneResult.event_count.participant.outdoors;
                                wallchartItem.Milestone1PersonalGrowth = (int)milestoneResult.event_count.participant.personal_growth;
                                wallchartItem.Milestone1Assist = PALEventCountTotal(milestoneResult.event_count.assistant);
                                wallchartItem.Milestone1Lead = PALEventCountTotal(milestoneResult.event_count.leader);
                            }
                            break;
                        case 2:
                            if (milestoneResult.status == "awarded")
                            {
                                wallchartItem.Milestone2Community = 5;
                                wallchartItem.Milestone2Creative = 5;
                                wallchartItem.Milestone2Outdoors = 5;
                                wallchartItem.Milestone2PersonalGrowth = 5;
                                wallchartItem.Milestone2Assist = 0;
                                wallchartItem.Milestone2Lead = 4;
                                wallchartItem.Milestone2Awarded = milestoneResult.status_updated;
                                wallchartItem.Milestone2Presented = GetPresentedDate(member.id, "Milestone 2");
                            }
                            else
                            {
                                wallchartItem.Milestone2Community = (int)milestoneResult.event_count.participant.community;
                                wallchartItem.Milestone2Creative = (int)milestoneResult.event_count.participant.creative;
                                wallchartItem.Milestone2Outdoors = (int)milestoneResult.event_count.participant.outdoors;
                                wallchartItem.Milestone2PersonalGrowth = (int)milestoneResult.event_count.participant.personal_growth;
                                wallchartItem.Milestone2Assist = PALEventCountTotal(milestoneResult.event_count.assistant);
                                wallchartItem.Milestone2Lead = PALEventCountTotal(milestoneResult.event_count.leader);
                            }
                            break; 
                        case 3:
                            wallchartItem.Milestone3Community = (int)milestoneResult.event_count.participant.community;
                            wallchartItem.Milestone3Creative = (int)milestoneResult.event_count.participant.creative;
                            wallchartItem.Milestone3Outdoors = (int)milestoneResult.event_count.participant.outdoors;
                            wallchartItem.Milestone3PersonalGrowth = (int)milestoneResult.event_count.participant.personal_growth;
                            wallchartItem.Milestone3Assist = PALEventCountTotal(milestoneResult.event_count.assistant);
                            wallchartItem.Milestone3Lead = PALEventCountTotal(milestoneResult.event_count.leader);
                            wallchartItem.Milestone3Awarded = milestoneResult.status == "awarded" ? milestoneResult.status_updated : null;
                            wallchartItem.Milestone3Presented = GetPresentedDate(member.id, "Milestone 3");
                            break;
                    }
                }
                // Get OAS for member
                var oasResults = await _terrainAPIService.GetOASResultsForMember(member.id);
                foreach (var oasResult in oasResults.results.Where(o => o.status == "awarded").OrderBy(o => o.achievement_meta.stream).ThenBy(o => o.achievement_meta.stage))
                {
                    switch (oasResult.achievement_meta.stream)
                    {
                        case "bushcraft":
                            wallchartItem.OASBushcraftStage = oasResult.achievement_meta.stage;
                            break;
                        case "bushwalking":
                            wallchartItem.OASBushwalkingStage = oasResult.achievement_meta.stage;
                            break;
                        case "camping":
                            wallchartItem.OASCampingStage = oasResult.achievement_meta.stage;
                            break;
                        case "alpine":
                            wallchartItem.OASAlpineStage = oasResult.achievement_meta.stage;
                            break;
                        case "cycling":
                            wallchartItem.OASCyclingStage = oasResult.achievement_meta.stage;
                            break;
                        case "vertical":
                            wallchartItem.OASVerticalStage = oasResult.achievement_meta.stage;
                            break;
                        case "aquatics":
                            wallchartItem.OASAquaticsStage = oasResult.achievement_meta.stage;
                            break;
                        case "boating":
                            wallchartItem.OASBoatingStage = oasResult.achievement_meta.stage;
                            break;
                        case "paddling":
                            wallchartItem.OASPaddlingStage = oasResult.achievement_meta.stage;
                            break;
                    }
                }
                wallchartItem.OASStageProgressions = AchievementSelection.OasProgressions(oasResults.results, _storageService.Section, includeApproved: false);
                // Member
                wallchartItem.MemberId = member.id;
                wallchartItem.MemberName = $"{member.first_name} {_membersService.GetMemberLastName(selectedUnitId, member.id).Result}";
                wallchartItem.MemberPatrol = member.patrol_name;
                wallchartItem.IntroToScouting = await GetIntroToScouting(member.id);
                wallchartItem.IntroToSection = await GetIntroToSection(member.id);
                wallchartItem.YearsOfService = YearsOfServiceCalculator.CalculateCompletedYears(wallchartItem.IntroToSection, DateTime.Now);
                // Logbook
                await _terrainAPIService.AssumeProfile(member.id);
                var getMemberLogbookMetrics = await _terrainAPIService.GetMemberLogbookMetrics(member.id);
                await _terrainAPIService.RevokeAssumedProfiles();
                var totalNightsCamped = getMemberLogbookMetrics.results.Where(r => r.name == "total_nights_camped").FirstOrDefault()?.value ?? 0;
                var totalNightsCampedInSection = totalNightsCamped;
                var totalKmsHiked = (getMemberLogbookMetrics.results.Where(r => r.name == "total_distance_hiked").FirstOrDefault()?.value ?? 0) / 1000.0f;
                wallchartItem.NightsCamped = (int)totalNightsCamped;
                wallchartItem.NightsCampedInSection = (int)totalNightsCampedInSection;
                wallchartItem.KMsHiked = (int)totalKmsHiked;
                // SIA
                var siaResultModel = await _terrainAPIService.GetSIAResultsForMember(member.id);
                var awardedSIAProjects = siaResultModel.results.Where(r => r.section == section && r.status == "awarded");
                var awardedSIAProjectsByArea = awardedSIAProjects.GroupBy(p => p.answers.special_interest_area_selection);
                foreach (var area in awardedSIAProjectsByArea)
                {
                    switch (area.Key)
                    {
                        case "sia_adventure_sport":
                            wallchartItem.SIAAdventureSport = area.Count();
                            break;
                        case "sia_art_literature":
                            wallchartItem.SIAArtsLiterature = area.Count();
                            break;
                        case "sia_environment":
                            wallchartItem.SIAEnvironment = area.Count();
                            break;
                        case "sia_stem_innovation":
                            wallchartItem.SIAStemInnovation = area.Count();
                            break;
                        case "sia_growth_development":
                            wallchartItem.SIAGrowthDevelopment = area.Count();
                            break;
                        case "sia_better_world":
                            wallchartItem.SIACreatingABetterWorld = area.Count();
                            break;
                    }
                }
                // Peak Award
                var courseReflectionResultModel = await _terrainAPIService.GetCourseReflectionResultsForMember(member.id);
                wallchartItem.LeadershipCourse = AchievementSelection.LatestAwardedPeak(courseReflectionResultModel.results, _storageService.Section)?.status_updated;

                var adventurousJourneyResultModel = await _terrainAPIService.GetAdventurousJourneyResultsForMember(member.id);
                wallchartItem.AdventurousJourney = AchievementSelection.LatestAwardedPeak(adventurousJourneyResultModel.results, _storageService.Section)?.status_updated;

                var personalReflectionResultModel = await _terrainAPIService.GetPersonalReflectionResultsForMember(member.id);
                wallchartItem.PersonalReflection = AchievementSelection.LatestAwardedPeak(personalReflectionResultModel.results, _storageService.Section)?.status_updated;

                var peakAwardResultModel = await _terrainAPIService.GetPeakAwardResultsForMember(member.id);
                wallchartItem.PeakAwardAwarded = AchievementSelection.LatestAwardedPeak(peakAwardResultModel.results, _storageService.Section)?.status_updated;

                // Terrain's own percentage, as on its Group Life page, for display until the award itself is granted (see PeakAwardAwarded).
                // Terrain shows whole percents, so round the same way to show the same figure.
                wallchartItem.PeakAward = terrainPeakProgress.TryGetValue(member.id, out var terrainPercent)
                    ? Math.Round(Math.Clamp(terrainPercent, 0, 100), 0, MidpointRounding.AwayFromZero) / 100.0
                    : 0;

                wallchartItems.Add(wallchartItem);
            }
            _storageService.CachedWallchartItems.Add(new KeyValuePair<string, List<WallchartItemModel>>(selectedUnitId, wallchartItems));
            return wallchartItems;
        }

        private void MoveMilestoneAssistsToLeads(MilestoneResult milestoneResult)
        {
            milestoneResult.event_count.leader.community += milestoneResult.event_count.assistant.community;
            milestoneResult.event_count.leader.outdoors += milestoneResult.event_count.assistant.outdoors;
            milestoneResult.event_count.leader.creative += milestoneResult.event_count.assistant.creative;
            milestoneResult.event_count.leader.personal_growth += milestoneResult.event_count.assistant.personal_growth;
            milestoneResult.event_count.assistant.community = 0;
            milestoneResult.event_count.assistant.outdoors = 0;
            milestoneResult.event_count.assistant.creative = 0;
            milestoneResult.event_count.assistant.personal_growth = 0;
        }

        private async Task<DateTime?> GetIntroToScouting(string memberId)
        {
            var introToScouting = await _terrainAPIService.GetIntroductionToScoutingResultsForMember(memberId);
            var result = AchievementSelection.LatestAwardedIntro(introToScouting.results);
            if (result != null)
            {
                if (!string.IsNullOrEmpty(result.imported.date_awarded))
                {
                    return ConvertStringToDate(result.imported.date_awarded, "yyyy-MM-dd");
                }
                else
                {
                    return result.status_updated;
                }
            }

            return null;
        }

        private async Task<DateTime?> GetIntroToSection(string memberId)
        {
            var introToSection = await _terrainAPIService.GetIntroductionToSectionResultsForMember(memberId);
            var result = AchievementSelection.LatestAwardedIntro(introToSection.results, _storageService.Section);
            if (result != null)
            {
                if (!string.IsNullOrEmpty(result.imported.date_awarded))
                {
                    return ConvertStringToDate(result.imported.date_awarded, "yyyy-MM-dd");
                }
                else
                {
                    return result.status_updated;
                }
            }

            return null;
        }

        private DateTime ConvertStringToDate(string dateString, string dateFormat)
        {
            return DateTime.ParseExact(dateString, dateFormat, CultureInfo.InvariantCulture);
        }

        private DateTime? GetPresentedDate(string member_id, string achievement_name)
        {
            var approval = approvals.Where(a => a.member_id == member_id && a.achievement_name == achievement_name).FirstOrDefault();
            return approval?.presented_date;
        }

        private int PALEventCountTotal(PALEventCount eventCount)
        {
            return (int)(eventCount.community + eventCount.creative + eventCount.outdoors + eventCount.personal_growth);
        }
    }
}
