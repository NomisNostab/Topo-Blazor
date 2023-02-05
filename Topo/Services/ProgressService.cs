using Newtonsoft.Json;
using System.Globalization;
using System.Reflection;
using Topo.Model.Members;
using Topo.Model.OAS;
using Topo.Model.Progress;

namespace Topo.Services
{
    public interface IProgressService
    {
        public Task<ProgressDetailsPageViewModel> GetProgressDetailsPageViewModel(string memberId);
    }

    public class ProgressService : IProgressService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly IMembersService _memberService;
        private int _oasProgressionCount = 0;

        public ProgressService(StorageService storageService, ITerrainAPIService terrainAPIService, IMembersService memberService)
        {
            _storageService = storageService;
            _terrainAPIService = terrainAPIService;
            _memberService = memberService;
        }

        public async Task<ProgressDetailsPageViewModel> GetProgressDetailsPageViewModel(string memberId)
        {
            ProgressDetailsPageViewModel progressDetailsPageViewModel = new ProgressDetailsPageViewModel();
            var members = await _memberService.GetMembersAsync(_storageService.UnitId);
            progressDetailsPageViewModel.Member = members.Where(m => m.id == memberId).FirstOrDefault() ?? new MemberListModel();

            progressDetailsPageViewModel.IntroToScoutingDate = await GetIntroToScouting(memberId);
            progressDetailsPageViewModel.IntroToSectionDate = await GetIntroToSection(memberId);
            progressDetailsPageViewModel.Milestones = await GetMilestoneSummaries(memberId);
            progressDetailsPageViewModel.OASSummaries = await GetOASSummaries(memberId);
            await _terrainAPIService.RevokeAssumedProfiles();
            await _terrainAPIService.AssumeProfile(memberId);
            progressDetailsPageViewModel.Stats = await GetStats(memberId);
            await _terrainAPIService.RevokeAssumedProfiles();
            progressDetailsPageViewModel.SIASummaries = await GetSIASummaries(memberId);
            progressDetailsPageViewModel.PeakAward = await GetPeakAwardSummary(memberId);

            return progressDetailsPageViewModel;
        }

        private async Task<string> GetIntroToScouting(string memberId)
        {
            var introToScoutingDate = "";
            var introToScouting = await _terrainAPIService.GetIntroductionToScoutingResultsForMember(memberId);
            var result = introToScouting.results.Where(r => r.status == "awarded").FirstOrDefault();
            if (result != null)
            {
                if (!string.IsNullOrEmpty(result.imported.date_awarded))
                {
                    introToScoutingDate = ConvertStringToDate(result.imported.date_awarded, "yyyy-MM-dd").ToString("dd/MM/yy");
                }
                else
                {
                    introToScoutingDate = result.status_updated.ToString("dd/MM/yy");
                }
            }

            return introToScoutingDate;
        }

        private async Task<string> GetIntroToSection(string memberId)
        {
            var introToSectionDate = "";
            var introToSection = await _terrainAPIService.GetIntroductionToSectionResultsForMember(memberId);
            var result = introToSection.results.Where(r => r.section == _storageService.Section && r.status == "awarded").FirstOrDefault();
            if (result != null)
            {
                if (!string.IsNullOrEmpty(result.imported.date_awarded))
                {
                    introToSectionDate = ConvertStringToDate(result.imported.date_awarded, "yyyy-MM-dd").ToString("dd/MM/yy");
                }
                else
                {
                    introToSectionDate = result.status_updated.ToString("dd/MM/yy");
                }
            }

            return introToSectionDate;
        }

        private DateTime ConvertStringToDate(string dateString, string dateFormat)
        {
            return DateTime.ParseExact(dateString, dateFormat, CultureInfo.InvariantCulture);
        }

        private async Task<List<MilestoneSummary>> GetMilestoneSummaries(string memberId)
        {
            var milestoneSummaries = new List<MilestoneSummary>();
            var milestones = await _terrainAPIService.GetMilestoneResultsForMember(memberId);
            foreach (var milestoneResult in milestones.results.Where(r => r.section == _storageService.Section).OrderBy(r => r.achievement_meta.stage))
            {
                var participantCount = 0;
                var assistCount = 0;
                var leadCount = 0;
                switch (milestoneResult.achievement_meta.stage)
                {
                    case 1:
                        participantCount = 6;
                        assistCount = 2;
                        leadCount = 1;
                        break;
                    case 2:
                        participantCount = 5;
                        assistCount = 3;
                        leadCount = 2;
                        break;
                    case 3:
                        participantCount = 4;
                        assistCount = 4;
                        leadCount = 4;
                        break;
                }
                var milestoneSummary = new MilestoneSummary();
                milestoneSummary.Milestone = milestoneResult.achievement_meta.stage;
                switch (milestoneResult.status)
                {
                    case "awarded":
                        milestoneSummary.Awarded = string.IsNullOrEmpty(milestoneResult.imported.date_awarded)
                            ? milestoneResult.status_updated
                            : ConvertStringToDate(milestoneResult.imported.date_awarded, "yyyy-MM-dd");
                        milestoneSummary.Status = "Awarded";
                        break;
                    case "not_required":
                        milestoneSummary.Status = "Not Required";
                        break;
                    case "in_progress":
                    default:
                        milestoneSummary.Status = "In Progress";
                        break;
                }
                if (milestoneResult.status == "awarded")
                {
                }

                // Add dummy imported entries
                if (milestoneResult.imported != null)
                {
                    var challengeArea = "";
                    var importedParticipantCount = 0;
                    var importedAssistCount = 0;
                    var importedLeadCount = 0;
                    for (int c = 0; c < 4; c++)
                    {

                        switch (c)
                        {
                            case 0:
                                challengeArea = "community";
                                importedParticipantCount = (int)milestoneResult.imported.event_count.participant.community;
                                importedAssistCount = (int)milestoneResult.imported.event_count.assistant.community;
                                importedLeadCount = (int)milestoneResult.imported.event_count.leader.community;
                                break;
                            case 1:
                                challengeArea = "creative";
                                importedParticipantCount = (int)milestoneResult.imported.event_count.participant.creative;
                                importedAssistCount = (int)milestoneResult.imported.event_count.assistant.creative;
                                importedLeadCount = (int)milestoneResult.imported.event_count.leader.creative;
                                break;
                            case 2:
                                challengeArea = "personal_growth";
                                importedParticipantCount = (int)milestoneResult.imported.event_count.participant.personal_growth;
                                importedAssistCount = (int)milestoneResult.imported.event_count.assistant.personal_growth;
                                importedLeadCount = (int)milestoneResult.imported.event_count.leader.personal_growth;
                                break;
                            case 3:
                                challengeArea = "outdoors";
                                importedParticipantCount = (int)milestoneResult.imported.event_count.participant.outdoors;
                                importedAssistCount = (int)milestoneResult.imported.event_count.assistant.outdoors;
                                importedLeadCount = (int)milestoneResult.imported.event_count.leader.outdoors;
                                break;
                        }
                        var importedDate = string.IsNullOrEmpty(milestoneResult.imported.date_awarded) 
                            ? milestoneResult.status_updated 
                            : ConvertStringToDate(milestoneResult.imported.date_awarded, "yyyy-MM-dd");
                        for (int i = 0; i < importedParticipantCount; i++)
                        {
                            milestoneSummary.ParticipateLogs.Add(
                                new MilestoneLog
                                {
                                    ChallengeArea = challengeArea,
                                    EventName = "Imported",
                                    EventDate = importedDate
                                });
                        }

                        for (int i = 0; i < importedAssistCount; i++)
                        {
                            milestoneSummary.AssistLogs.Add(
                                new MilestoneLog
                                {
                                    ChallengeArea = challengeArea,
                                    EventName = "Imported",
                                    EventDate = importedDate
                                });
                        }

                        for (int i = 0; i < importedLeadCount; i++)
                        {
                            milestoneSummary.LeadLogs.Add(
                                new MilestoneLog
                                {
                                    ChallengeArea = challengeArea,
                                    EventName = "Imported",
                                    EventDate = importedDate
                                });
                        }
                    }
                }

                var participantEvents = milestoneResult.event_log.Where(e => e.credit_type == "participant").ToList();
                var groupedLogs = participantEvents.GroupBy(e => e.challenge_area);
                foreach (var groupedLog in groupedLogs)
                {

                    foreach (var milestoneEvent in groupedLog.Take(participantCount))
                        milestoneSummary.ParticipateLogs.Add(
                            new MilestoneLog
                            {
                                ChallengeArea = milestoneEvent.challenge_area,
                                EventName = milestoneEvent.event_name,
                                EventDate = milestoneEvent.event_start_datetime
                            });
                }

                var assistEvents = milestoneResult.event_log.Where(e => e.credit_type == "assistant").ToList();
                foreach (var assistEvent in assistEvents.Take(assistCount))
                {
                    milestoneSummary.AssistLogs.Add(
                        new MilestoneLog
                        {
                            ChallengeArea = assistEvent.challenge_area,
                            EventName = assistEvent.event_name,
                            EventDate = assistEvent.event_start_datetime
                        });
                }

                var leadEvents = milestoneResult.event_log.Where(e => e.credit_type == "leader").ToList();
                foreach (var leadEvent in leadEvents.Take(leadCount))
                {
                    milestoneSummary.LeadLogs.Add(
                        new MilestoneLog
                        {
                            ChallengeArea = leadEvent.challenge_area,
                            EventName = leadEvent.event_name,
                            EventDate = leadEvent.event_start_datetime
                        });
                }

                milestoneSummaries.Add(milestoneSummary);
            }

            return milestoneSummaries;
        }

        private async Task<List<OASSummary>> GetOASSummaries(string memberId)
        {
            var oasSummaries = new List<OASSummary>();

            var oasResults = await _terrainAPIService.GetOASResultsForMember(memberId);
            foreach (var oasResult in oasResults.results.OrderBy(o => o.achievement_meta.stream).ThenBy(o => o.achievement_meta.stage).ThenBy(o => o.status_updated))
            {
                var awardedDate = new DateTime();
                if (oasResult.status == "awarded")
                {
                    awardedDate = oasResult.status_updated;
                    if (!string.IsNullOrEmpty(oasResult.imported?.date_awarded))
                    {
                        awardedDate = ConvertStringToDate(oasResult.imported?.date_awarded, "yyyy-MM-dd");
                    }
                }
                oasSummaries.Add(new OASSummary
                {
                    Stream = oasResult.achievement_meta.stream,
                    Stage = oasResult.achievement_meta.stage,
                    Awarded = awardedDate
                });
            }

            _oasProgressionCount = oasResults.results.Where(r => r.status == "awarded" && r.section == _storageService.Section).Count();
            return oasSummaries;
        }

        private async Task<Stats> GetStats(string memberId)
        {
            var stats = new Stats();

            var getMemberLogbookMetrics = await _terrainAPIService.GetMemberLogbookMetrics(memberId);
            var totalNightsCamped = getMemberLogbookMetrics.results.Where(r => r.name == "total_nights_camped").FirstOrDefault()?.value ?? 0;
            var totalKmsHiked = (getMemberLogbookMetrics.results.Where(r => r.name == "total_distance_hiked").FirstOrDefault()?.value ?? 0) / 1000.0f;

            stats.OasProgressions = _oasProgressionCount;
            stats.KmsHiked = (int)totalKmsHiked;
            stats.NightsCamped = (int)totalNightsCamped;

            return stats;
        }

        private async Task<List<SIASummary>> GetSIASummaries(string memberId)
        {
            TextInfo myTI = new CultureInfo("en-UA", false).TextInfo;

            var siaResultModel = await _terrainAPIService.GetSIAResultsForMember(memberId);
            var memberSiaProjects = siaResultModel.results.Where(r => r.section == _storageService.Section)
                .Select(r => new SIASummary
                {
                    Area = myTI.ToTitleCase(r.answers.special_interest_area_selection.Replace("sia_", "").Replace("_", " ")),
                    Project = r.answers.project_name,
                    Status = myTI.ToTitleCase(r.status.Replace("_", " ")),
                    Awarded = r.status_updated
                });

            return memberSiaProjects.ToList();
        }

        private async Task<PeakAward> GetPeakAwardSummary(string memberid)
        {
            var peakAward = new PeakAward();

            var courseReflectionResultModel = await _terrainAPIService.GetCourseReflectionResultsForMember(memberid);
            peakAward.PersonalDevelopmentCourse = courseReflectionResultModel.results.Where(r => r.status == "awarded").FirstOrDefault()?.status_updated.ToString("dd/MM/yy") ?? "";

            var adventurousJourneyResultModel = await _terrainAPIService.GetAdventurousJourneyResultsForMember(memberid);
            peakAward.AdventurousJourney = adventurousJourneyResultModel.results.Where(r => r.status == "awarded").FirstOrDefault()?.status_updated.ToString("dd/MM/yy") ?? "";

            var personalReflectionResultModel = await _terrainAPIService.GetPersonalReflectionResultsForMember(memberid);
            peakAward.PersonalReflection = personalReflectionResultModel.results.Where(r => r.status == "awarded").FirstOrDefault()?.status_updated.ToString("dd/MM/yy") ?? "";

            var peakAwardResultModel = await _terrainAPIService.GetPeakAwardResultsForMember(memberid);
            peakAward.Awarded = peakAwardResultModel.results.Where(r => r.status == "awarded").FirstOrDefault()?.status_updated.ToString("dd/MM/yy") ?? "";

            return peakAward;
        }
    }
}
