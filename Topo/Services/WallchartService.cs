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
        Task<List<WallchartItemModel>> GetWallchartItems(string selectedUnitId);
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

        public async Task<List<WallchartItemModel>> GetWallchartItems(string selectedUnitId)
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

            var wallchartItems = new List<WallchartItemModel>();
            foreach (var member in members.Where(m => m.isAdultLeader == 0).OrderBy(m => m.last_name).ThenBy(m => m.first_name))
            {
                var wallchartItem = new WallchartItemModel();

                // Get Milestone for member
                var memberMilestones = await _terrainAPIService.GetMilestoneResultsForMember(member.id);
                foreach (var milestoneResult in memberMilestones.results.Where(r => r.section == _storageService.Section).OrderBy(r => r.achievement_meta.stage))
                {
                    MoveMilestoneAssistsToLeads(milestoneResult);
                    switch (milestoneResult.achievement_meta.stage)
                    {
                        case 1:
                            wallchartItem.Milestone1Skipped = milestoneResult.status == "not_required";
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
                            wallchartItem.Milestone2Skipped = milestoneResult.status == "not_required";
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
                            wallchartItem.Milestone3Skipped = milestoneResult.status == "not_required";
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
                wallchartItem.OASStageProgressions = oasResults.results.Where(o => o.status == "awarded" && o.section == _storageService.Section).Count();
                // Double Count OAS Stage Progressions for stages 7 and above.
                wallchartItem.OASStageProgressions += oasResults.results.Where(o => o.achievement_meta.stage >=7 && o.status == "awarded" && o.section == _storageService.Section).Count();
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
                wallchartItem.LeadershipCourse = courseReflectionResultModel.results
                    .Where(r => r.status == "awarded" && r.section == _storageService.Section)
                    .FirstOrDefault()?.status_updated;

                var adventurousJourneyResultModel = await _terrainAPIService.GetAdventurousJourneyResultsForMember(member.id);
                wallchartItem.AdventurousJourney = adventurousJourneyResultModel.results
                    .Where(r => r.status == "awarded" && r.section == _storageService.Section)
                    .FirstOrDefault()?.status_updated;

                var personalReflectionResultModel = await _terrainAPIService.GetPersonalReflectionResultsForMember(member.id);
                wallchartItem.PersonalReflection = personalReflectionResultModel.results
                    .Where(r => r.status == "awarded" && r.section == _storageService.Section)
                    .FirstOrDefault()?.status_updated;

                var peakAwardResultModel = await _terrainAPIService.GetPeakAwardResultsForMember(member.id);
                wallchartItem.PeakAwardAwarded = peakAwardResultModel.results
                    .Where(r => r.status == "awarded" && r.section == _storageService.Section)
                    .FirstOrDefault()?.status_updated;

                // Percent complete, for display until the award itself is granted (see PeakAwardAwarded).
                wallchartItem.PeakAward = ComputePeakAwardPercent(wallchartItem, _storageService.Section);

                wallchartItems.Add(wallchartItem);
            }
            _storageService.CachedWallchartItems.Add(new KeyValuePair<string, List<WallchartItemModel>>(selectedUnitId, wallchartItems));
            return wallchartItems;
        }

        // Peak Award percent complete = an equal-weighted average of a few "buckets", each itself a
        // 0.0-1.0 fraction: Milestones, OAS, SIA, and the three individual awards (Leadership Course,
        // Adventurous Journey, Personal Reflection). See PeakAwardRequirements for the per-section
        // numeric thresholds this depends on.
        private static double ComputePeakAwardPercent(WallchartItemModel w, string section)
        {
            var buckets = new List<double>
            {
                ComputeMilestoneBucket(w),
                w.LeadershipCourse.HasValue ? 1.0 : 0.0,
                w.AdventurousJourney.HasValue ? 1.0 : 0.0,
                w.PersonalReflection.HasValue ? 1.0 : 0.0,
            };

            // OAS and SIA need section-specific thresholds; if the section isn't one we have
            // requirements for, those two buckets are simply left out rather than guessed at.
            if (PeakAwardRequirements.BySection.TryGetValue(section, out var requirement))
            {
                buckets.Add(ComputeOASBucket(w, requirement));
                buckets.Add(ComputeSIABucket(w, requirement));
            }

            return buckets.Average();
        }

        // Each milestone contributes its own percent-complete (same formula as the Milestone report)
        // rather than a flat awarded/not-awarded count, so partial progress on the current milestone
        // shows up here too. A milestone that's been awarded or marked "not required" always counts
        // as fully done - the raw counts behind an awarded milestone don't necessarily add up to the
        // formula's own targets (the award is a leader's call, not a strict numeric gate), and a
        // skipped milestone's counts are 0 regardless of progress.
        private static double ComputeMilestoneBucket(WallchartItemModel w)
        {
            var milestone1 = w.Milestone1Awarded.HasValue || w.Milestone1Skipped
                ? 1.0
                : MilestonePercentComplete(1, w.Milestone1Community, w.Milestone1Outdoors, w.Milestone1Creative, w.Milestone1PersonalGrowth, w.Milestone1Lead);
            var milestone2 = w.Milestone2Awarded.HasValue || w.Milestone2Skipped
                ? 1.0
                : MilestonePercentComplete(2, w.Milestone2Community, w.Milestone2Outdoors, w.Milestone2Creative, w.Milestone2PersonalGrowth, w.Milestone2Lead);
            var milestone3 = w.Milestone3Awarded.HasValue || w.Milestone3Skipped
                ? 1.0
                : MilestonePercentComplete(3, w.Milestone3Community, w.Milestone3Outdoors, w.Milestone3Creative, w.Milestone3PersonalGrowth, w.Milestone3Lead);
            return new[] { milestone1, milestone2, milestone3 }.Average();
        }

        private static double MilestonePercentComplete(int level, int community, int outdoors, int creative, int personalGrowth, int lead)
        {
            return Math.Min(1.0, MilestoneProgressCalculator.CalculatePercentComplete(level, community, outdoors, creative, personalGrowth, lead) / 100.0);
        }

        private static double ComputeOASBucket(WallchartItemModel w, PeakAwardRequirements.SectionRequirement requirement)
        {
            var ratios = new List<double>
            {
                SafeRatio(w.OASStageProgressions, requirement.OASProgressionsRequired),
                SafeRatio(w.OASBushcraftStage, requirement.OASCoreMinLevel),
                SafeRatio(w.OASBushwalkingStage, requirement.OASCoreMinLevel),
                SafeRatio(w.OASCampingStage, requirement.OASCoreMinLevel),
            };

            if (requirement.OASLevel4PlusRequired > 0)
            {
                var streamsAtLevel4Plus = new[]
                {
                    w.OASBushcraftStage, w.OASBushwalkingStage, w.OASCampingStage,
                    w.OASAlpineStage, w.OASCyclingStage, w.OASVerticalStage,
                    w.OASAquaticsStage, w.OASBoatingStage, w.OASPaddlingStage
                }.Count(stage => stage >= 4);
                ratios.Add(SafeRatio(streamsAtLevel4Plus, requirement.OASLevel4PlusRequired));
            }

            return ratios.Average();
        }

        private static double ComputeSIABucket(WallchartItemModel w, PeakAwardRequirements.SectionRequirement requirement)
        {
            var siaCountsByArea = new[]
            {
                w.SIAAdventureSport, w.SIAArtsLiterature, w.SIAEnvironment,
                w.SIAStemInnovation, w.SIAGrowthDevelopment, w.SIACreatingABetterWorld
            };
            var totalRatio = SafeRatio(siaCountsByArea.Sum(), requirement.SIARequired);
            var areaSpreadRatio = SafeRatio(siaCountsByArea.Count(count => count > 0), requirement.SIAMinAreasRequired);
            return new[] { totalRatio, areaSpreadRatio }.Average();
        }

        // A requirement of 0 (or a negative typo) means "nothing needed", i.e. trivially 100% met -
        // without this, PeakAwardRequirements.cs being hand-edited to a 0 would silently turn the
        // whole Peak Award percentage into NaN (0/0) instead of a clear, harmless 100%.
        private static double SafeRatio(double achieved, int required)
        {
            if (required <= 0)
                return 1.0;
            return Math.Min(1.0, achieved / required);
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
            var result = introToScouting.results.Where(r => r.status == "awarded").FirstOrDefault();
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
            var result = introToSection.results.Where(r => r.section == _storageService.Section && r.status == "awarded").FirstOrDefault();
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
