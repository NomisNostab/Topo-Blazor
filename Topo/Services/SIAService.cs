using System.Globalization;
using Topo.Model.Approvals;
using Topo.Model.SIA;

namespace Topo.Services
{
    public interface ISIAService
    {
        public Task<List<SIAProjectListModel>> GenerateSIAReportData(List<KeyValuePair<string, string>> selectedMembers, string section, string selectedUnitId);
    }

    public class SIAService : ISIAService
    {
        private readonly ITerrainAPIService _terrainAPIService;

        public SIAService(ITerrainAPIService terrainAPIService)
        {
            _terrainAPIService = terrainAPIService;
        }

        public async Task<List<SIAProjectListModel>> GenerateSIAReportData(List<KeyValuePair<string, string>> selectedMembers, string section, string selectedUnitId)
        {
            var approvals = new List<ApprovalsListModel>(); // await _approvalsService.GetApprovalListItems(selectedUnitId);
            TextInfo myTI = new CultureInfo("en-UA", false).TextInfo;
            var unitSiaProjects = new List<SIAProjectListModel>();
            foreach (var member in selectedMembers)
            {
                try
                {
                    var siaResultModel = await _terrainAPIService.GetSIAResultsForMember(member.Key);
                    var memberSiaProjects = siaResultModel.results.Where(r => r.section == section)
                        .Select(r => new SIAProjectListModel
                        {
                            memberName = member.Value,
                            area = myTI.ToTitleCase(r.answers.special_interest_area_selection.Replace("sia_", "").Replace("_", " ")),
                            projectName = r.answers.project_name,
                            status = myTI.ToTitleCase(r.status.Replace("_", " ")),
                            statusUpdated = r.status_updated,
                            achievement_id = r.id
                        })
                        .ToList();
                    foreach (var memberProject in memberSiaProjects)
                    {
                        var approval = approvals.Where(a => a.achievement_id == memberProject.achievement_id && a.submission_type.ToLower() == "review").FirstOrDefault();
                        if (approval != null)
                        {
                            if (approval.presented_date.HasValue)
                            {
                                memberProject.status = "Presented";
                                memberProject.statusUpdated = approval.presented_date.Value.ToLocalTime();
                            }
                        }
                    }
                    if (memberSiaProjects != null && memberSiaProjects.Count > 0)
                        unitSiaProjects = unitSiaProjects.Concat(memberSiaProjects).ToList();
                }
                catch (Exception ex)
                {
                    unitSiaProjects.Add(new SIAProjectListModel
                    {
                        memberName = $"{member.Value}",
                        area = "Error",
                        projectName = "Error",
                        status = "Error",
                        statusUpdated = DateTime.Now
                    });
                }
            }

            return unitSiaProjects;
        }

    }
}
