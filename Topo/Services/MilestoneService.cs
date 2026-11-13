using Newtonsoft.Json;
using Topo.Model.Milestone;

namespace Topo.Services
{
    public interface IMilestoneService
    {
        public Task<List<MilestoneSummaryListModel>> GetMilestoneSummaries(string selectedUnitId, Action<string>? progress = null);
        // Names the summaries again, for when the last names are shown or hidden after they were read
        public Task RenameMembers(string selectedUnitId, List<MilestoneSummaryListModel> summaries);
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

        public async Task<List<MilestoneSummaryListModel>> GetMilestoneSummaries(string selectedUnitId, Action<string>? progress = null)
        {
            var unitMilestoneSummary = new List<MilestoneSummaryListModel>();
            // Read the unit's milestones once instead of once per youth. Every record carries its member_id, so ToLookup hands the loop
            // a member's records as before (a member with none gets an empty list). The records include earlier sections, filtered below.
            var milestonesByMember = (await _terrainAPIService.GetUnitAchievementsOfType<GetMilestoneResultModel>(selectedUnitId, "milestone")).results.ToLookup(r => r.member_id);
            var members = (await _membersService.GetMembersAsync(selectedUnitId)).Where(m => m.isAdultLeader == 0).ToList();
            for (var i = 0; i < members.Count; i++)
            {
                var member = members[i];
                progress?.Invoke($"{member.first_name} {member.last_name} ({i + 1} of {members.Count})");
                MilestoneResult milestone1 = new MilestoneResult();
                MilestoneResult milestone2 = new MilestoneResult();
                MilestoneResult milestone3 = new MilestoneResult();
                // "Awarded" here also covers Terrain's "approved" status: once a leader has approved
                // a milestone it's effectively done (awarding is just the formal presentation step
                // that follows later), so a member should move on to progress at the next milestone
                // as soon as the current one is approved, not only once it's been formally awarded.
                // See IsAwardedOrApproved - the same "awarded" + "approved" pairing is already used
                // for OAS elsewhere in this app (OASService/ProgressService).
                bool milestone1Awarded = false;
                bool milestone2Awarded = false;
                bool milestone3Awarded = false;
                bool milestone1Skipped = true;
                bool milestone2Skipped = true;
                bool milestone3Skipped = false;
                int currentLevel = 1;
                foreach (var milestoneResult in milestonesByMember[member.id].Where(r => r.section == _storageService.Section).OrderBy(r => r.achievement_meta.stage))
                {
                    MoveMilestoneAssistsToLeads(milestoneResult);
                    switch (milestoneResult.achievement_meta.stage)
                    {
                        case 1:
                            milestone1 = milestoneResult;
                            milestone1Awarded = IsAwardedOrApproved(milestone1.status);
                            milestone1Skipped = milestone1.status == "not_required";
                            currentLevel = (milestone1Awarded || milestone1Skipped) ? 2 : 1;
                            break;
                        case 2:
                            milestone2 = milestoneResult;
                            milestone2Awarded = IsAwardedOrApproved(milestone2.status);
                            milestone2Skipped = milestone2.status == "not_required";
                            currentLevel = (milestone1Awarded || milestone1Skipped) ? 2 : 1;
                            currentLevel = (milestone1Awarded || milestone1Skipped) && (milestone2Awarded || milestone2Skipped) ? 3 : currentLevel;
                            break;
                        case 3:
                            milestone3 = milestoneResult;
                            milestone3Awarded = IsAwardedOrApproved(milestone3.status);
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
                        // Unlike milestones 1 and 2, there's no "Milestone 4" to move on to once this
                        // is awarded, so the raw event-count formula below would otherwise keep showing
                        // a member's progress as it stood at the point of award, even though the report
                        // now lists them as having completed it (see milestone3Awarded above).
                        percentComplete = milestone3Awarded ? 100 : CalculateMilestonePercentComplete(3, milestone3.event_count);
                        break;
                }
                unitMilestoneSummary.Add(
                    new MilestoneSummaryListModel
                    {
                        memberId = member.id,
                        memberName = await MemberName(selectedUnitId, member),
                        memberPatrol = member.patrol_name,
                        currentLevel = currentLevel,
                        percentComplete = percentComplete,
                        milestone3Awarded = milestone3Awarded,
                        milestone1ParticipateCommunity = milestone1Skipped ? 0 : (milestone1Awarded ? 6 : (int)milestone1.event_count.participant.community),
                        milestone1ParticipateOutdoors = milestone1Skipped ? 0 : (milestone1Awarded ? 6 : (int)milestone1.event_count.participant.outdoors),
                        milestone1ParticipateCreative = milestone1Skipped ? 0 : (milestone1Awarded ? 6 : (int)milestone1.event_count.participant.creative),
                        milestone1ParticipatePersonalGrowth = milestone1Skipped ? 0 : (milestone1Awarded ? 6 : (int)milestone1.event_count.participant.personal_growth),
                        milestone1Assist = milestone1Skipped ? 0 : (milestone1Awarded ? 2 : (int)milestone1.event_count.assistant.community 
                                                                                            + (int)milestone1.event_count.assistant.outdoors
                                                                                            + (int)milestone1.event_count.assistant.creative
                                                                                            + (int)milestone1.event_count.assistant.personal_growth),
                        milestone1AssistArea = milestone1Skipped ? "" : FormatMilestoneChallengeArea(milestone1.event_count.assistant),
                        milestone1Lead = milestone1Skipped ? 0 : (milestone1Awarded ? 2 : (int)milestone1.event_count.leader.community
                                                                                            + (int)milestone1.event_count.leader.outdoors
                                                                                            + (int)milestone1.event_count.leader.creative
                                                                                            + (int)milestone1.event_count.leader.personal_growth),
                        milestone1LeadArea = milestone1Skipped ? "" : FormatMilestoneChallengeArea(milestone1.event_count.leader),
                        milestone2ParticipateCommunity = milestone2Skipped ? 0 : (milestone2Awarded ? 5 : (int)milestone2.event_count.participant.community),
                        milestone2ParticipateOutdoors = milestone2Skipped ? 0 : (milestone2Awarded ? 5 : (int)milestone2.event_count.participant.outdoors),
                        milestone2ParticipateCreative = milestone2Skipped ? 0 : (milestone2Awarded ? 5 : (int)milestone2.event_count.participant.creative),
                        milestone2ParticipatePersonalGrowth = milestone2Skipped ? 0 : (milestone2Awarded ? 5 : (int)milestone2.event_count.participant.personal_growth),
                        milestone2Assist = milestone2Skipped ? 0 : (milestone2Awarded ? 3 : (int)milestone2.event_count.assistant.community
                                                                                            + (int)milestone2.event_count.assistant.outdoors
                                                                                            + (int)milestone2.event_count.assistant.creative
                                                                                            + (int)milestone2.event_count.assistant.personal_growth),
                        milestone2AssistArea = milestone2Skipped ? "" : FormatMilestoneChallengeArea(milestone2.event_count.assistant),
                        milestone2Lead = milestone2Skipped ? 0 : (milestone2Awarded ? 4 : (int)milestone2.event_count.leader.community
                                                                                            + (int)milestone2.event_count.leader.outdoors
                                                                                            + (int)milestone2.event_count.leader.creative
                                                                                            + (int)milestone2.event_count.leader.personal_growth),
                        milestone2LeadArea = milestone2Skipped ? "" : FormatMilestoneChallengeArea(milestone2.event_count.leader),
                        milestone3ParticipateCommunity = milestone3Skipped ? 0 : (milestone3Awarded ? 4 : (int)milestone3.event_count.participant.community),
                        milestone3ParticipateOutdoors = milestone3Skipped ? 0 : (milestone3Awarded ? 4 : (int)milestone3.event_count.participant.outdoors),
                        milestone3ParticipateCreative = milestone3Skipped ? 0 : (milestone3Awarded ? 4 : (int)milestone3.event_count.participant.creative),
                        milestone3ParticipatePersonalGrowth = milestone3Skipped ? 0 : (milestone3Awarded ? 4 : (int)milestone3.event_count.participant.personal_growth),
                        milestone3Assist = milestone3Awarded ? 4 : (int)milestone3.event_count.assistant.community
                                                                + (int)milestone3.event_count.assistant.outdoors
                                                                + (int)milestone3.event_count.assistant.creative
                                                                + (int)milestone3.event_count.assistant.personal_growth,
                        milestone3AssistArea = milestone3Skipped ? "" : FormatMilestoneChallengeArea(milestone3.event_count.assistant),
                        milestone3Lead = milestone3Awarded ? 6 : (int)milestone3.event_count.leader.community
                                                                + (int)milestone3.event_count.leader.outdoors
                                                                + (int)milestone3.event_count.leader.creative
                                                                + (int)milestone3.event_count.leader.personal_growth,
                        milestone3LeadArea = milestone3Skipped ? "" : FormatMilestoneChallengeArea(milestone3.event_count.leader)
                    });
            }
            return unitMilestoneSummary;
        }

        public async Task RenameMembers(string selectedUnitId, List<MilestoneSummaryListModel> summaries)
        {
            var members = await _membersService.GetMembersAsync(selectedUnitId);
            foreach (var summary in summaries)
            {
                var member = members.FirstOrDefault(m => m.id == summary.memberId);
                if (member != null)
                    summary.memberName = await MemberName(selectedUnitId, member);
            }
        }

        // The first name and as much of the last as the Suppress Last Name setting allows
        private async Task<string> MemberName(string selectedUnitId, Topo.Model.Members.MemberListModel member) =>
            $"{member.first_name} {await _membersService.GetMemberLastName(selectedUnitId, member.id)}";

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

        private string FormatMilestoneChallengeArea(PALEventCount eventCount)
        {
            var community = (int)eventCount.community == 1 ? "Co " : (int)eventCount.community > 1 ? $"Co-{(int)eventCount.community} " : "";
            var outdoors = (int)eventCount.outdoors == 1 ? "O " : (int)eventCount.outdoors > 1 ? $"O-{(int)eventCount.outdoors} " : "";
            var creative = (int)eventCount.creative == 1 ? "Cr " : (int)eventCount.creative > 1 ? $"Cr-{(int)eventCount.creative} " : "";
            var personalGrowth = (int)eventCount.personal_growth == 1 ? "PG " : (int)eventCount.personal_growth > 1 ? $"PG-{(int)eventCount.personal_growth} " : "";
            return  (community + outdoors + creative + personalGrowth).TrimEnd();
        }

        private int CalculateMilestonePercentComplete(int currentLevel, Event_Count eventCount)
        {
            int target = 0;
            int participantTotal = (int)((currentLevel == 3 ? Math.Min(4.0, eventCount.participant.community) : eventCount.participant.community)
                                + (currentLevel == 3 ? Math.Min(4.0, eventCount.participant.outdoors) : eventCount.participant.outdoors)
                                + (currentLevel == 3 ? Math.Min(4.0, eventCount.participant.creative) : eventCount.participant.creative)
                                + (currentLevel == 3 ? Math.Min(4.0, eventCount.participant.personal_growth) : eventCount.participant.personal_growth));
            int leaderTotal = (int)(eventCount.leader.community
                                 + eventCount.leader.creative
                                 + eventCount.leader.outdoors
                                 + eventCount.leader.personal_growth);
            var total = currentLevel == 3
                    ? participantTotal + Math.Min(6, leaderTotal)
                    : participantTotal + leaderTotal;

            switch (currentLevel)
            {
                case 1:
                    target = 26;
                    break;
                case 2:
                    target = 24;
                    break;
                case 3:
                    target = 20;
                    break;
            }

            var percentComplete = total * 100 / target;
            return percentComplete;
        }

    }
}
