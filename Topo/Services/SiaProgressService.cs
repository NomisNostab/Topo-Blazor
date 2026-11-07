using Topo.Model.Members;
using Topo.Model.SiaProgress;

namespace Topo.Services
{
    public class SiaProgressRequest
    {
        public SiaScope Scope { get; set; }
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        // The unit's youth members
        public List<MemberListModel> Members { get; set; } = new List<MemberListModel>();
        // For one project or one person
        public string MemberId { get; set; } = string.Empty;
        public string ProjectId { get; set; } = string.Empty;
        public bool ExcludeCompleted { get; set; }
        public bool IncludeSummary { get; set; }
        public bool DoubleSided { get; set; }
    }

    public interface ISiaProgressService
    {
        public Task<List<SiaProjectChoice>> GetProjectsAsync(string memberId, string section);
        public Task<SiaProgressReport> GetReportAsync(SiaProgressRequest request, Action<string>? progress = null);
    }

    public class SiaProgressService : ISiaProgressService
    {
        private readonly ITerrainAPIService _terrainAPIService;
        // The questions are the same for everyone in a section, so they are read once
        private readonly Dictionary<string, SiaTemplates> _templates = new Dictionary<string, SiaTemplates>();

        public SiaProgressService(ITerrainAPIService terrainAPIService)
        {
            _terrainAPIService = terrainAPIService;
        }

        public async Task<List<SiaProjectChoice>> GetProjectsAsync(string memberId, string section)
        {
            var results = await _terrainAPIService.GetSIAResultsForMember(memberId);
            return results.results
                .Where(r => string.Equals(r.section, section, StringComparison.OrdinalIgnoreCase))
                .OrderBy(r => r.status_updated)
                .Select(r => new SiaProjectChoice
                {
                    Id = r.id,
                    Label = $"{(string.IsNullOrWhiteSpace(r.answers.project_name) ? "(no name yet)" : r.answers.project_name)} - {Status(r.status)}"
                })
                .ToList();
        }

        public async Task<SiaProgressReport> GetReportAsync(SiaProgressRequest request, Action<string>? progress = null)
        {
            var members = request.Scope == SiaScope.WholeUnit
                ? SiaProgressBuilder.InPatrolOrder(request.Members)
                : request.Members.Where(m => m.id == request.MemberId).ToList();

            var input = new SiaProgressInput
            {
                UnitName = request.UnitName,
                Section = request.Section,
                ProjectId = request.Scope == SiaScope.OneProject ? request.ProjectId : "",
                ExcludeCompleted = request.ExcludeCompleted,
                IncludeSummary = request.IncludeSummary,
                DoubleSided = request.DoubleSided,
                Templates = await GetTemplatesAsync(request.Section)
            };
            if (PeakAwardRequirements.BySection.TryGetValue(request.Section, out var requirement))
            {
                input.AwardName = requirement.AwardName;
                input.SiaRequired = requirement.SIARequired;
                input.AreasRequired = requirement.SIAMinAreasRequired;
            }

            for (var i = 0; i < members.Count; i++)
            {
                progress?.Invoke($"{members[i].first_name} {members[i].last_name} ({i + 1} of {members.Count})");
                input.Members.Add(new SiaMemberInput
                {
                    MemberId = members[i].id,
                    MemberName = $"{members[i].first_name} {members[i].last_name}".Trim(),
                    Results = await _terrainAPIService.GetSIAResultsForMember(members[i].id)
                });
            }

            return SiaProgressBuilder.Build(input);
        }

        private async Task<SiaTemplates> GetTemplatesAsync(string section)
        {
            if (!_templates.TryGetValue(section, out var templates))
            {
                templates = new SiaTemplates
                {
                    Plan = SiaTemplateParser.Parse(await _terrainAPIService.GetSIATemplateJsonAsync(section, "plan")),
                    Review = SiaTemplateParser.Parse(await _terrainAPIService.GetSIATemplateJsonAsync(section, "review"))
                };
                // a template that couldn't be read is tried again next time
                if (templates.Plan.Questions.Count > 0)
                    _templates[section] = templates;
            }
            return templates;
        }

        private static string Status(string status)
        {
            var text = (status ?? "").Replace("_", " ");
            return text.Length == 0 ? "" : char.ToUpperInvariant(text[0]) + text.Substring(1);
        }
    }
}
