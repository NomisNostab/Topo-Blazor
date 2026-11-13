using Topo.Model.Members;
using Topo.Model.SIA;
using Topo.Model.SiaProgress;

namespace Topo.Services
{
    public class SiaProgressRequest
    {
        public SiaScope Scope { get; set; }
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public string UnitId { get; set; } = string.Empty;
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

            if (request.Scope == SiaScope.WholeUnit)
            {
                // One read gives every Scout's projects, so the unit isn't read once per Scout. A unit with no youth has nothing to read.
                ILookup<string, Result>? byMember = null;
                if (members.Count > 0)
                {
                    progress?.Invoke("the unit's projects");
                    byMember = (await _terrainAPIService.GetUnitAchievementsOfType<GetSIAResultsModel>(request.UnitId, "special_interest_area"))
                        .results.ToLookup(r => r.member_id);
                }
                foreach (var member in members)
                    input.Members.Add(MemberInput(member, new GetSIAResultsModel { results = byMember?[member.id].ToArray() ?? new Result[0] }));
                return input;
            }

            // one Scout: the Reading line names them the short way, as the reports do
            var shortNames = SiaProgressBuilder.ShortNames(members);
            for (var i = 0; i < members.Count; i++)
            {
                progress?.Invoke($"{(shortNames.TryGetValue(members[i].id, out var shortName) ? shortName : members[i].first_name)} ({i + 1} of {members.Count})");
                input.Members.Add(MemberInput(members[i], await _terrainAPIService.GetSIAResultsForMember(members[i].id)));
            }

            return input;
        }

        private static SiaMemberInput MemberInput(MemberListModel member, GetSIAResultsModel results) => new SiaMemberInput
        {
            MemberId = member.id,
            MemberName = $"{member.first_name} {member.last_name}".Trim(),
            Results = results
        };

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
