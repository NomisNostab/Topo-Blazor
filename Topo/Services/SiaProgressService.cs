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
    }

    public interface ISiaProgressService
    {
        // Reads what the report needs from Terrain. The report is built from it with SiaProgressBuilder, so its options can
        // change without reading again.
        public Task<SiaProgressInput> LoadAsync(SiaProgressRequest request, Action<string>? progress = null);
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

        public async Task<SiaProgressInput> LoadAsync(SiaProgressRequest request, Action<string>? progress = null)
        {
            var members = request.Scope == SiaScope.WholeUnit
                ? SiaProgressBuilder.InPatrolOrder(request.Members)
                : request.Members.Where(m => m.id == request.MemberId).ToList();

            var input = new SiaProgressInput
            {
                UnitName = request.UnitName,
                Section = request.Section,
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

            return input;
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
    }
}
