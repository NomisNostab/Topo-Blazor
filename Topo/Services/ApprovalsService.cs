using Blazored.LocalStorage;
using Microsoft.AspNetCore.Components.Forms;
using Newtonsoft.Json;
using System.Globalization;
using Topo.Model.Approvals;
using Topo.Model.Members;

namespace Topo.Services
{
    public interface IApprovalsService
    {
        public Task<List<ApprovalsListModel>> GetApprovalListItems(string unitId);
        public Task<List<ApprovalsListModel>> ReadApprovalListFromLocalStorage(string unitId);
        public Task UpdateApproval(string unitId, ApprovalsListModel approval);
        public Task<string> DownloadApprovalList(string unitId);
        public Task UploadApprovals(IBrowserFile approvalsFile, string unitId);
    }
    public class ApprovalsService : IApprovalsService
    {
        private readonly IMembersService _memberService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly ILocalStorageService _localStorageService;

        public ApprovalsService(IMembersService membersService, ITerrainAPIService terrainAPIService, ILocalStorageService localStorageService)
        {
            _memberService = membersService;
            _terrainAPIService = terrainAPIService;
            _localStorageService = localStorageService;
        }

        //private async Task<List<ApprovalsListModel>> GetApprovalList(string unitId, string status)
        //{
        //    TextInfo myTI = new CultureInfo("en-AU", false).TextInfo;
        //    var members = await _memberService.GetMembersAsync(unitId);
        //    var approvalList = await _terrainAPIService.GetUnitApprovals(unitId, status);
        //    var approvals = new List<ApprovalsListModel>();
        //    foreach (var approval in approvalList.results)
        //    {
        //        var member = members.Where(m => m.id == approval.member.id).FirstOrDefault();
        //        if (member != null)
        //        {
        //            DateTime? awardedDate;
        //            switch (approval.submission.type)
        //            {
        //                case "review":
        //                    awardedDate = (approval.submission.status == "finalised" && approval.submission.outcome == "approved") ? approval.submission.date : null;
        //                    break;
        //                case "award":
        //                    if (approval.submission.actioned_by.Length != 0)
        //                    {
        //                        //awardedDate = approval.submission.actioned_by.OrderBy(m => m.id).FirstOrDefault().;
        //                    }
        //                    break;
        //                default:
        //                    break;
        //            }

        //            approvals.Add(new ApprovalsListModel()
        //            {
        //                member_id = approval.member.id,
        //                member_first_name = approval.member.first_name,
        //                member_last_name = approval.member.last_name,
        //                achievement_id = approval.achievement.id,
        //                achievement_type = approval.achievement.type,
        //                submission_type = myTI.ToTitleCase(approval.submission.type),
        //                submission_status = myTI.ToTitleCase(approval.submission.status),
        //                submission_outcome = myTI.ToTitleCase(approval.submission.outcome),
        //                submission_date = approval.submission.date,
        //                awarded_date = (approval.submission.status == "finalised" && approval.submission.outcome == "approved") ? approval.submission.date : null
        //            });
        //        }
        //    }

        //    return approvals;
        //}

        //private async Task<List<ApprovalsListModel>> GetAdditionalAwardList(string unitId)
        //{
        //    TextInfo myTI = new CultureInfo("en-AU", false).TextInfo;
        //    var members = await _memberService.GetMembersAsync(unitId);
        //    var unitAchievementsResult = await _terrainAPIService.GetUnitAdditionalAwardAchievements(unitId);
        //    var approvals = new List<ApprovalsListModel>();
        //    foreach (var additionalAward in unitAchievementsResult.results)
        //    {
        //        var member = members.Where(m => m.id == additionalAward.member_id).FirstOrDefault();
        //        if (member != null)
        //        {
        //            var awardStatusDate = additionalAward.status_updated;
        //            if (additionalAward.imported != null)
        //                awardStatusDate = DateTime.ParseExact(additionalAward.imported.date_awarded, "yyyy-MM-dd", CultureInfo.InvariantCulture);
        //            approvals.Add(new ApprovalsListModel()
        //            {
        //                member_id = additionalAward.member_id,
        //                member_first_name = member?.first_name ?? "",
        //                member_last_name = member?.last_name ?? "",
        //                achievement_id = additionalAward.id,
        //                achievement_type = additionalAward.achievement_meta.additional_award_id,
        //                submission_type = "Review",
        //                submission_status = "Finalised",
        //                submission_outcome = myTI.ToTitleCase(additionalAward.status),
        //                submission_date = awardStatusDate,
        //                awarded_date = awardStatusDate
        //            });
        //        }
        //    }

        //    return approvals;
        //}

        private async Task WriteApprovalsListToLocalStarage(List<ApprovalsListModel> approvalsList, string unitId)
        {
            await _localStorageService.SetItemAsync($"{unitId}_ApprovalsList", approvalsList);
        }

        //private async Task<string> GetAchievementName(string member_id, string achievement_id, string achievement_type)
        //{
        //    var name = "";
        //    switch (achievement_type)
        //    {
        //        case "intro_scouting":
        //            name = "intro to scouting";
        //            break;
        //        case "intro_section":
        //            name = "intro to section";
        //            break;
        //        case "milestone":
        //            var memberMilestoneAchievementResult = await _terrainAPIService.GetMemberAchievementResult(member_id, achievement_id, achievement_type);
        //            name = memberMilestoneAchievementResult != null ? $"milestone {memberMilestoneAchievementResult.achievement_meta.stage}" : "milestone not found";
        //            break;
        //        case "outdoor_adventure_skill":
        //            var memberAchievementResult = await _terrainAPIService.GetMemberAchievementResult(member_id, achievement_id, achievement_type);
        //            if (memberAchievementResult != null && !string.IsNullOrEmpty(memberAchievementResult.template))
        //            {
        //                var templateParts = memberAchievementResult.template.Split("/");
        //                name = string.Join(" ", templateParts);
        //            }
        //            else
        //            {
        //                name = "outdoor adventure skill not found";
        //            }
        //            break;
        //        case "special_interest_area":
        //            var siaResult = await _terrainAPIService.GetSIAResultForMember(member_id, achievement_id);
        //            if (siaResult != null && siaResult.answers != null && !string.IsNullOrEmpty(siaResult.answers.special_interest_area_selection) && !string.IsNullOrEmpty(siaResult.answers.project_name))
        //                name = $"{siaResult.answers.special_interest_area_selection.Replace("_", " ")} {siaResult.answers.project_name}";
        //            else
        //                name = "SIA Not Found";
        //            break;
        //        default:
        //            name = achievement_type.Replace("_", " ");
        //            break;
        //    }
        //    TextInfo myTI = new CultureInfo("en-AU", false).TextInfo;
        //    return myTI.ToTitleCase(name);
        //}

        public async Task<List<ApprovalsListModel>> GetApprovalListItems(string unitId)
        {
            var pendingApprovals = await _terrainAPIService.GetUnitApprovals(unitId, "pending");
            var finalisedApprovals = await _terrainAPIService.GetUnitApprovals(unitId, "finalised");
            var approvals = finalisedApprovals.results.Concat(pendingApprovals.results).ToList();
            var achievementApprovals = approvals.GroupBy(r => r.achievement.id).ToList();
            var memberList = await _memberService.GetMembersAsync(unitId);
            List<ApprovalsListModel> approvalList = new List<ApprovalsListModel>();
            foreach (var result in achievementApprovals)
            {
                string approvalDate = "";
                string awardedDate = "";
                ApprovalsListModel approvalItem = new ApprovalsListModel();
                foreach (var item in result)
                {
                    Console.Write($"Type:{item.achievement.type} ");
                    approvalItem.achievement_type = item.achievement.type;
                    approvalItem.achievement_id = item.achievement.id;
                    switch (item.achievement.type)
                    {
                        case "special_interest_area":
                            var sia_area = string.IsNullOrEmpty(item.achievement.achievement_meta?.sia_area) ? "special_interest_area" : item.achievement.achievement_meta?.sia_area;
                            Console.Write($"Area:{sia_area} ");
                            approvalItem.achievement_name = $"{sia_area.Replace("_", " ")}";
                            break;
                        case "milestone":
                            Console.Write($"Stage:{item.achievement.achievement_meta?.stage} ");
                            approvalItem.achievement_name = $"milestone {item.achievement.achievement_meta?.stage}";
                            break;
                        case "outdoor_adventure_skill":
                            Console.Write($"Stream:{item.achievement.achievement_meta?.stream} Stage:{item.achievement.achievement_meta?.stage} ");
                            approvalItem.achievement_name = item.achievement.achievement_meta?.stream == item.achievement.achievement_meta?.branch
                                    ? $"OAS {item.achievement.achievement_meta?.stream} {item.achievement.achievement_meta?.stage}"
                                    : $"OAS {item.achievement.achievement_meta?.stream} {item.achievement.achievement_meta?.branch} {item.achievement.achievement_meta?.stage}";
                            break;
                        case "intro_section":
                            approvalItem.achievement_name = "intro to section";
                            break;
                        case "intro_scouting":
                            approvalItem.achievement_name = "intro to scouting";
                            break;
                        default:
                            approvalItem.achievement_name = item.achievement.type.Replace("_", " ");
                            break;
                    }
                    TextInfo myTI = new CultureInfo("en-AU", false).TextInfo;
                    approvalItem.achievement_name = myTI.ToTitleCase(approvalItem.achievement_name);

                    approvalItem.member_id = item.member.id;
                    approvalItem.member_first_name = item.member.first_name;
                    approvalItem.member_last_name = item.member.last_name;
                    Console.Write($"Name:{item.member.first_name} {item.member.last_name} ");

                    approvalItem.submission_type = myTI.ToTitleCase(item.submission.type);
                    if (item.submission.type == "review")
                    {
                        approvalItem.submission_date = item.submission.date;
                    }
                    approvalItem.submission_status = myTI.ToTitleCase(item.submission.status);
                    approvalItem.submission_outcome = myTI.ToTitleCase(item.submission.outcome);

                    Console.Write($"Submission Type:{item.submission.type} Status:{item.submission.status} Outcome:{item.submission.outcome} ");
                    int actionCount = 0;
                    foreach (var action in item.submission.actioned_by)
                    {
                        if (action.outcome == "approved")
                        {
                            actionCount++;
                        }
                        Console.Write($"Action outcome:{action.outcome} Action time:{action.time} Action date:{action.date_awarded} ");
                        switch (item.submission.type)
                        {
                            case "review":
                                if (actionCount == 2)
                                {
                                    approvalDate = action.time.ToShortDateString();
                                    approvalItem.awarded_date = action.time;
                                }
                                break;
                            case "award":
                                // need to check if action.member_id is a leader
                                var actionMember = memberList.Where(m => m.id == action.member_id).FirstOrDefault();
                                bool isLeader = actionMember != null && actionMember.isAdultLeader == 1;
                                if (isLeader)
                                {
                                    awardedDate = string.IsNullOrEmpty(action.date_awarded) ? action.time.ToShortDateString() : action.date_awarded;
                                    approvalItem.presented_date = action.time;
                                    if (!string.IsNullOrEmpty(action.date_awarded))
                                    {
                                        approvalItem.presented_date = DateTime.ParseExact(action.date_awarded, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                                    }
                                }
                                break;
                        }
                    }
                    Console.WriteLine();
                }
                Console.WriteLine($"Approved: {approvalDate} Awarded:{awardedDate}");
                approvalList.Add(approvalItem);
            }
            return approvalList;
        }
        //public async Task<List<ApprovalsListModel>> GetApprovalListItemsx(string unitId)
        //{
        //    var savedApprovalItems = await ReadApprovalListFromLocalStorage(unitId);
        //    var members = await _memberService.GetMembersAsync(unitId);
        //    List<ApprovalsListModel> approvalsToRemove = new List<ApprovalsListModel>();
        //    foreach (var approvalItem in savedApprovalItems)
        //    {
        //        var member = members.Where(m => m.id == approvalItem.member_id).FirstOrDefault();
        //        if (member == null)
        //        {
        //            approvalsToRemove.Add(approvalItem);
        //        }
        //    }
        //    foreach (var approvalToRemove in approvalsToRemove)
        //    {
        //        savedApprovalItems.Remove(approvalToRemove);
        //    }
        //    var initialLoad = savedApprovalItems.Count == 0;
        //    var pendingApprovals = await GetApprovalList(unitId, "pending");
        //    var finalisedApprovals = await GetApprovalList(unitId, "finalised");
        //    var additionalAwards = await GetAdditionalAwardList(unitId);
        //    var allTerrainApprovals = finalisedApprovals.Concat(pendingApprovals).Concat(additionalAwards).OrderBy(a => a.submission_date);
        //    // Remove pending approvals from savedApprovalItems
        //    var oldPendingApprovalItems = savedApprovalItems.Where(s => s.submission_status.ToLower() == "pending").ToList();
        //    if (oldPendingApprovalItems != null && oldPendingApprovalItems.Any())
        //    {
        //        foreach (var pendingItem in oldPendingApprovalItems)
        //        {
        //            savedApprovalItems.Remove(pendingItem);
        //        }
        //    }
        //    // Get items in allTerrainApprovals that are not in savedApprovalItems, these are new since last time
        //    var newSubmissions = allTerrainApprovals.Where(all => savedApprovalItems.Count(x => x.achievement_id == all.achievement_id && x.submission_type == all.submission_type) == 0).ToList();

        //    foreach (var newApproval in newSubmissions)
        //    {
        //        newApproval.achievement_name = await GetAchievementName(newApproval.member_id, newApproval.achievement_id, newApproval.achievement_type);
        //        if (initialLoad)
        //            newApproval.presented_date = newApproval.awarded_date;
        //    }

        //    savedApprovalItems.AddRange(newSubmissions);

        //    await WriteApprovalsListToLocalStarage(savedApprovalItems.OrderBy(a => a.submission_date).ToList(), unitId);

        //    return savedApprovalItems ?? new List<ApprovalsListModel>();
        //}

        public async Task<List<ApprovalsListModel>> ReadApprovalListFromLocalStorage(string unitId)
        {
            List<ApprovalsListModel>? list = await _localStorageService.GetItemAsync<List<ApprovalsListModel>>($"{unitId}_ApprovalsList");

            return list ?? new List<ApprovalsListModel>();
        }

        public async Task UpdateApproval(string unitId, ApprovalsListModel approval)
        {
            var savedApprovalItems = await ReadApprovalListFromLocalStorage(unitId);
            var approvalItem = savedApprovalItems.Where(a => a.achievement_id == approval.achievement_id && a.submission_type == approval.submission_type).FirstOrDefault();
            if (approvalItem != null)
            {
                approvalItem.presented_date = approval.presented_date.HasValue ? approval.presented_date.Value.ToLocalTime() : null;
            }
            await WriteApprovalsListToLocalStarage(savedApprovalItems.OrderBy(a => a.submission_date).ToList(), unitId);
        }

        public async Task<string> DownloadApprovalList(string unitId)
        {
            var savedApprovalItems = await ReadApprovalListFromLocalStorage(unitId); 
            return JsonConvert.SerializeObject(savedApprovalItems);
        }

        public async Task UploadApprovals(IBrowserFile approvalsFile, string unitId)
        {
            try
            {
                var json = await new StreamReader(approvalsFile.OpenReadStream()).ReadToEndAsync();
                List<ApprovalsListModel>? list = JsonConvert.DeserializeObject<List<ApprovalsListModel>>(json);
                await WriteApprovalsListToLocalStarage(list?.OrderBy(a => a.submission_date).ToList() ?? new List<ApprovalsListModel>(), unitId);
            }
            catch
            {
                return;
            }
        }
    }
}
