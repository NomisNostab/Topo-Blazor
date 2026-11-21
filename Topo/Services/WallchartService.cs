using Topo.Model;
using Topo.Model.Approvals;
using Topo.Model.Milestone;
using Topo.Model.Members;
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

            var youth = members.Where(m => m.isAdultLeader == 0).OrderBy(m => m.last_name).ThenBy(m => m.first_name).ToList();
            if (youth.Count == 0)
            {
                // a unit with no youth (an adults-only team) has nothing to show, and nothing to read
                var none = new List<WallchartItemModel>();
                _storageService.CachedWallchartItems.Add(new KeyValuePair<string, List<WallchartItemModel>>(selectedUnitId, none));
                return none;
            }

            // Everything the wallchart reads from Terrain is asked for at once instead of one read after another:
            //  - the unit-wide reads, none of which depends on another or needs a profile taken on, and
            //  - each Scout's logbook totals, which can only be read one Scout at a time (Terrain keeps one profile taken on at a
            //    time), but which don't stop the unit-wide reads from working, so the two overlap instead of waiting for each other.
            progress?.Invoke("the unit's records");
            var logbookRead = ReadLogbookTotals(youth, progress);
            // Terrain works out each member's Peak Award Progress itself; read it rather than estimating it here
            var peakProgressRead = _terrainAPIService.GetUnitMetrics(selectedUnitId);
            // Each kind of record is read once for the whole unit instead of once per member.
            // Every record carries its member_id, so ToLookup lets the loop below ask for "this member's records" as it did before;
            // a member with none gets an empty list.
            var milestonesRead = _terrainAPIService.GetUnitAchievementsOfType<GetMilestoneResultModel>(selectedUnitId, "milestone");
            var oasRead = _terrainAPIService.GetUnitAchievementsOfType<Topo.Model.OAS.GetUnitAchievementsResultsModel>(selectedUnitId, "outdoor_adventure_skill");
            var introScoutingRead = _terrainAPIService.GetUnitAchievementsOfType<GetIntroductionResultsModel>(selectedUnitId, "intro_scouting");
            var introSectionRead = _terrainAPIService.GetUnitAchievementsOfType<GetIntroductionResultsModel>(selectedUnitId, "intro_section");
            var siaRead = _terrainAPIService.GetUnitAchievementsOfType<Topo.Model.SIA.GetSIAResultsModel>(selectedUnitId, "special_interest_area");
            var courseReflectionRead = _terrainAPIService.GetUnitAchievementsOfType<GetPeakAwardsResultsModel>(selectedUnitId, "course_reflection");
            var adventurousJourneyRead = _terrainAPIService.GetUnitAchievementsOfType<GetPeakAwardsResultsModel>(selectedUnitId, "adventurous_journey");
            var personalReflectionRead = _terrainAPIService.GetUnitAchievementsOfType<GetPeakAwardsResultsModel>(selectedUnitId, "personal_reflection");
            var peakAwardRead = _terrainAPIService.GetUnitAchievementsOfType<GetPeakAwardsResultsModel>(selectedUnitId, "peak_award");
            await Task.WhenAll(logbookRead, peakProgressRead, milestonesRead, oasRead, introScoutingRead, introSectionRead, siaRead,
                courseReflectionRead, adventurousJourneyRead, personalReflectionRead, peakAwardRead);

            var logbookTotals = logbookRead.Result;
            var terrainPeakProgress = peakProgressRead.Result
                .GroupBy(m => m.member_id)
                .ToDictionary(g => g.Key, g => g.First().achievement_metrics.peak_award_progress);
            var milestonesByMember = milestonesRead.Result.results.ToLookup(r => r.member_id);
            var oasByMember = oasRead.Result.results.ToLookup(r => r.member_id);
            var introScoutingByMember = introScoutingRead.Result.results.ToLookup(r => r.member_id);
            var introSectionByMember = introSectionRead.Result.results.ToLookup(r => r.member_id);
            var siaByMember = siaRead.Result.results.ToLookup(r => r.member_id);
            var courseReflectionByMember = courseReflectionRead.Result.results.ToLookup(r => r.member_id);
            var adventurousJourneyByMember = adventurousJourneyRead.Result.results.ToLookup(r => r.member_id);
            var personalReflectionByMember = personalReflectionRead.Result.results.ToLookup(r => r.member_id);
            var peakAwardByMember = peakAwardRead.Result.results.ToLookup(r => r.member_id);

            var wallchartItems = new List<WallchartItemModel>();
            foreach (var member in youth)
            {
                var wallchartItem = new WallchartItemModel();

                // Get Milestone for member
                foreach (var milestoneResult in AchievementSelection.OnePerMilestoneStage(milestonesByMember[member.id].Where(r => r.section == _storageService.Section)))
                {
                    MoveMilestoneAssistsToLeads(milestoneResult);
                    switch (milestoneResult.achievement_meta.stage)
                    {
                        case 1:
                            if (IsAwardedOrApproved(milestoneResult.status))
                            {
                                wallchartItem.Milestone1Community = 6;
                                wallchartItem.Milestone1Creative = 6;
                                wallchartItem.Milestone1Outdoors = 6;
                                wallchartItem.Milestone1PersonalGrowth = 6;
                                wallchartItem.Milestone1Assist = 0;
                                wallchartItem.Milestone1Lead = 2;
                                wallchartItem.Milestone1Awarded = MilestoneDates.AwardedDate(milestoneResult);
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
                            if (IsAwardedOrApproved(milestoneResult.status))
                            {
                                wallchartItem.Milestone2Community = 5;
                                wallchartItem.Milestone2Creative = 5;
                                wallchartItem.Milestone2Outdoors = 5;
                                wallchartItem.Milestone2PersonalGrowth = 5;
                                wallchartItem.Milestone2Assist = 0;
                                wallchartItem.Milestone2Lead = 4;
                                wallchartItem.Milestone2Awarded = MilestoneDates.AwardedDate(milestoneResult);
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
                            // "Awarded" here also covers Terrain's "approved" status - see IsAwardedOrApproved.
                            wallchartItem.Milestone3Awarded = IsAwardedOrApproved(milestoneResult.status) ? MilestoneDates.AwardedDate(milestoneResult) : null;
                            wallchartItem.Milestone3Presented = GetPresentedDate(member.id, "Milestone 3");
                            break;
                    }
                }
                // Get OAS for member
                var oasResults = oasByMember[member.id];
                foreach (var oasResult in oasResults.Where(o => o.status == "awarded").OrderBy(o => o.achievement_meta.stream).ThenBy(o => o.achievement_meta.stage))
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
                wallchartItem.OASStageProgressions = AchievementSelection.OasProgressions(oasResults, _storageService.Section, includeApproved: false);
                // Member
                wallchartItem.MemberId = member.id;
                wallchartItem.MemberName = $"{member.first_name} {_membersService.GetMemberLastName(selectedUnitId, member.id).Result}";
                wallchartItem.MemberPatrol = member.patrol_name;
                wallchartItem.IntroToScouting = IntroDate(AchievementSelection.LatestAwardedIntro(introScoutingByMember[member.id]));
                wallchartItem.IntroToSection = IntroDate(AchievementSelection.LatestAwardedIntro(introSectionByMember[member.id], _storageService.Section));
                wallchartItem.YearsOfService = YearsOfServiceCalculator.CalculateCompletedYears(wallchartItem.IntroToSection, DateTime.Now);
                // Logbook
                var (totalNightsCamped, totalKmsHiked) = logbookTotals[member.id];
                var totalNightsCampedInSection = totalNightsCamped;
                wallchartItem.NightsCamped = (int)totalNightsCamped;
                wallchartItem.NightsCampedInSection = (int)totalNightsCampedInSection;
                wallchartItem.KMsHiked = (int)totalKmsHiked;
                // SIA
                var awardedSIAProjects = siaByMember[member.id].Where(r => r.section == section && r.status == "awarded");
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
                wallchartItem.LeadershipCourse = PeakDate(AchievementSelection.LatestAwardedPeak(courseReflectionByMember[member.id], _storageService.Section));
                wallchartItem.AdventurousJourney = PeakDate(AchievementSelection.LatestAwardedPeak(adventurousJourneyByMember[member.id], _storageService.Section));
                wallchartItem.PersonalReflection = PeakDate(AchievementSelection.LatestAwardedPeak(personalReflectionByMember[member.id], _storageService.Section));
                wallchartItem.PeakAwardAwarded = PeakDate(AchievementSelection.LatestAwardedPeak(peakAwardByMember[member.id], _storageService.Section));

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

        // Each Scout's logbook totals. A Scout's profile has to be taken on to read them, and Terrain keeps only one profile
        // taken on at a time, so this goes one Scout at a time. It is started alongside the unit-wide reads, which work while a
        // profile is taken on, so it doesn't hold them up.
        private async Task<Dictionary<string, (float nights, float kms)>> ReadLogbookTotals(List<MemberListModel> youth, Action<string>? progress)
        {
            var totals = new Dictionary<string, (float nights, float kms)>();
            try
            {
                for (var i = 0; i < youth.Count; i++)
                {
                    var member = youth[i];
                    progress?.Invoke($"{member.first_name} {member.last_name} ({i + 1} of {youth.Count})");
                    await _terrainAPIService.AssumeProfile(member.id);
                    var metrics = await _terrainAPIService.GetMemberLogbookMetrics(member.id);
                    await _terrainAPIService.RevokeAssumedProfiles();
                    var nights = metrics.results.Where(r => r.name == "total_nights_camped").FirstOrDefault()?.value ?? 0;
                    var kms = (metrics.results.Where(r => r.name == "total_distance_hiked").FirstOrDefault()?.value ?? 0) / 1000.0f;
                    totals[member.id] = (nights, kms);
                }
            }
            catch
            {
                // a failed read mustn't leave a Scout's profile taken on
                await _terrainAPIService.RevokeAssumedProfiles();
                throw;
            }
            return totals;
        }

        // A leader approving a milestone is functionally "done" - the formal award is just the
        // presentation step that follows later. Same "awarded" + "approved" pairing already used
        // for OAS elsewhere in this app (OASService/ProgressService) and for the Milestone report
        // (MilestoneService.IsAwardedOrApproved).
        private static bool IsAwardedOrApproved(string status) => status == "awarded" || status == "approved";

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

        // The award date to show for an introduction record, or nothing when the member has no awarded one
        private static DateTime? IntroDate(Topo.Model.Progress.Result? record) =>
            record == null ? null : AwardDate.Resolve(record.date_awarded, record.imported?.date_awarded, record.status_updated);

        private static DateTime? PeakDate(PeakAwardResult? record) =>
            record == null ? null : AwardDate.Resolve(record.date_awarded, null, record.status_updated);

        private static DateTime MilestoneAwardedDate(MilestoneResult milestoneResult) =>
            AwardDate.Resolve(milestoneResult.date_awarded, milestoneResult.imported?.date_awarded, milestoneResult.status_updated);

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
