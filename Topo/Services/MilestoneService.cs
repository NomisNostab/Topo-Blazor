using Newtonsoft.Json;
using Topo.Model.Milestone;

namespace Topo.Services
{
    public interface IMilestoneService
    {
        public Task<List<MilestoneSummaryListModel>> GetMilestoneSummaries(string selectedUnitId);
    }

    public class MilestoneService : IMilestoneService
    {
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly IMembersService _membersService;
        private readonly StorageService _storageService;

        public MilestoneService(ITerrainAPIService terrainAPIService, IMembersService membersService, StorageService storageService)
        {
            _terrainAPIService = terrainAPIService;
            _membersService = membersService;
            _storageService = storageService;
        }

        public async Task<List<MilestoneSummaryListModel>> GetMilestoneSummaries(string selectedUnitId)
        {
            var unitMilestoneSummary = new List<MilestoneSummaryListModel>();
            var members = await _membersService.GetMembersAsync(selectedUnitId);
            foreach (var member in members.Where(m => m.isAdultLeader == 0))
            {
                var memberMilestones = await _terrainAPIService.GetMilestoneResultsForMember(member.id);
                MilestoneResult milestone1 = new MilestoneResult();
                MilestoneResult milestone2 = new MilestoneResult();
                MilestoneResult milestone3 = new MilestoneResult();
                bool milestone1Awarded = false;
                bool milestone2Awarded = false;
                bool milestone3Awarded = false;
                bool milestone1Skipped = true;
                bool milestone2Skipped = true;
                bool milestone3Skipped = false;
                int currentLevel = 1;
                foreach (var milestoneResult in memberMilestones.results.Where(r => r.section == _storageService.Section).OrderBy(r => r.achievement_meta.stage))
                {
                    switch (milestoneResult.achievement_meta.stage)
                    {
                        case 1:
                            milestone1 = milestoneResult;
                            milestone1Awarded = milestone1.status == "awarded";
                            milestone1Skipped = milestone1.status == "not_required";
                            currentLevel = (milestone1Awarded || milestone1Skipped) ? 2 : 1;
                            break;
                        case 2:
                            milestone2 = milestoneResult;
                            milestone2Awarded = milestone2.status == "awarded";
                            milestone2Skipped = milestone2.status == "not_required";
                            currentLevel = (milestone1Awarded || milestone1Skipped) ? 2 : 1;
                            currentLevel = (milestone1Awarded || milestone1Skipped) && (milestone2Awarded || milestone2Skipped) ? 3 : currentLevel;
                            break;
                        case 3:
                            milestone3 = milestoneResult;
                            milestone3Awarded = milestone3.status == "awarded";
                            milestone3Skipped = milestone3.status == "not_required";
                            currentLevel = (milestone1Awarded || milestone1Skipped) ? 2 : 1;
                            currentLevel = (milestone1Awarded || milestone1Skipped) && (milestone2Awarded || milestone2Skipped) ? 3 : currentLevel;
                            currentLevel = (milestone1Awarded || milestone1Skipped) && (milestone2Awarded || milestone2Skipped) && (milestone3Awarded || milestone3Skipped) ? 3 : currentLevel;
                            break;
                    }
                }
                int percentComplete = 0;
                switch (currentLevel)
                {
                    case 1:
                        percentComplete = CalculateMilestonePercentComplete(1, milestone1.event_count);
                        break;
                    case 2:
                        percentComplete = CalculateMilestonePercentComplete(2, milestone2.event_count);
                        break; 
                    case 3:
                        percentComplete = CalculateMilestonePercentComplete(3, milestone3.event_count);
                        break;
                }
                unitMilestoneSummary.Add(
                    new MilestoneSummaryListModel
                    {
                        memberName = $"{member.first_name} {member.last_name}",
                        currentLevel = currentLevel,
                        percentComplete = percentComplete,
                        milestone1ParticipateCommunity = milestone1Skipped ? 0 : (milestone1Awarded ? 6 : (int)milestone1.event_count.participant.community),
                        milestone1ParticipateOutdoors = milestone1Skipped ? 0 : (milestone1Awarded ? 6 : (int)milestone1.event_count.participant.outdoors),
                        milestone1ParticipateCreative = milestone1Skipped ? 0 : (milestone1Awarded ? 6 : (int)milestone1.event_count.participant.creative),
                        milestone1ParticipatePersonalGrowth = milestone1Skipped ? 0 : (milestone1Awarded ? 6 : (int)milestone1.event_count.participant.personal_growth),
                        milestone1Assist = milestone1Skipped ? 0 : (milestone1Awarded ? 2 : (int)milestone1.event_count.assistant.community 
                                                                                            + (int)milestone1.event_count.assistant.outdoors
                                                                                            + (int)milestone1.event_count.assistant.creative
                                                                                            + (int)milestone1.event_count.assistant.personal_growth),
                        milestone1Lead = milestone1Skipped ? 0 : (milestone1Awarded ? 1 : (int)milestone1.event_count.leader.community
                                                                                            + (int)milestone1.event_count.leader.outdoors
                                                                                            + (int)milestone1.event_count.leader.creative
                                                                                            + (int)milestone1.event_count.leader.personal_growth),
                        milestone2ParticipateCommunity = milestone2Skipped ? 0 : (milestone2Awarded ? 5 : (int)milestone2.event_count.participant.community),
                        milestone2ParticipateOutdoors = milestone2Skipped ? 0 : (milestone2Awarded ? 5 : (int)milestone2.event_count.participant.outdoors),
                        milestone2ParticipateCreative = milestone2Skipped ? 0 : (milestone2Awarded ? 5 : (int)milestone2.event_count.participant.creative),
                        milestone2ParticipatePersonalGrowth = milestone2Skipped ? 0 : (milestone2Awarded ? 5 : (int)milestone2.event_count.participant.personal_growth),
                        milestone2Assist = milestone2Skipped ? 0 : (milestone2Awarded ? 3 : (int)milestone2.event_count.assistant.community
                                                                                            + (int)milestone2.event_count.assistant.outdoors
                                                                                            + (int)milestone2.event_count.assistant.creative
                                                                                            + (int)milestone2.event_count.assistant.personal_growth),
                        milestone2Lead = milestone2Skipped ? 0 : (milestone2Awarded ? 2 : (int)milestone2.event_count.leader.community
                                                                                            + (int)milestone2.event_count.leader.outdoors
                                                                                            + (int)milestone2.event_count.leader.creative
                                                                                            + (int)milestone2.event_count.leader.personal_growth),
                        milestone3ParticipateCommunity = milestone3Skipped ? 0 : (milestone3Awarded ? 4 : (int)milestone3.event_count.participant.community),
                        milestone3ParticipateOutdoors = milestone3Skipped ? 0 : (milestone3Awarded ? 4 : (int)milestone3.event_count.participant.outdoors),
                        milestone3ParticipateCreative = milestone3Skipped ? 0 : (milestone3Awarded ? 4 : (int)milestone3.event_count.participant.creative),
                        milestone3ParticipatePersonalGrowth = milestone3Skipped ? 0 : (milestone3Awarded ? 4 : (int)milestone3.event_count.participant.personal_growth),
                        milestone3Assist = milestone3Awarded ? 4 : (int)milestone3.event_count.assistant.community
                                                                + (int)milestone3.event_count.assistant.outdoors
                                                                + (int)milestone3.event_count.assistant.creative
                                                                + (int)milestone3.event_count.assistant.personal_growth,
                        milestone3Lead = milestone3Awarded ? 4 : (int)milestone3.event_count.leader.community
                                                                + (int)milestone3.event_count.leader.outdoors
                                                                + (int)milestone3.event_count.leader.creative
                                                                + (int)milestone3.event_count.leader.personal_growth,
                    });
            }
            return unitMilestoneSummary;
        }

        private int CalculateMilestonePercentComplete(int currentLevel, Event_Count eventCount)
        {
            int target = 0;
            int participantTotal = (int)((currentLevel == 3 ? Math.Min(4.0, eventCount.participant.community) : eventCount.participant.community)
                                + (currentLevel == 3 ? Math.Min(4.0, eventCount.participant.outdoors) : eventCount.participant.outdoors)
                                + (currentLevel == 3 ? Math.Min(4.0, eventCount.participant.creative) : eventCount.participant.creative)
                                + (currentLevel == 3 ? Math.Min(4.0, eventCount.participant.personal_growth) : eventCount.participant.personal_growth));
            int assistantTotal = (int)(eventCount.assistant.community
                                 + eventCount.assistant.creative
                                 + eventCount.assistant.outdoors
                                 + eventCount.assistant.personal_growth);
            int leaderTotal = (int)(eventCount.leader.community
                                 + eventCount.leader.creative
                                 + eventCount.leader.outdoors
                                 + eventCount.leader.personal_growth);
            var total = currentLevel == 3
                    ? participantTotal + Math.Min(4, assistantTotal) + Math.Min(4, leaderTotal)
                    : participantTotal + assistantTotal + leaderTotal;

            switch (currentLevel)
            {
                case 1:
                    target = 27;
                    break;
                case 2:
                    target = 25;
                    break;
                case 3:
                    target = 24;
                    break;
            }

            var percentComplete = total * 100 / target;
            return percentComplete;
        }

    }
}
