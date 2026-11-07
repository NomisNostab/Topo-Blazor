using Topo.Model.Members;
using Topo.Model.ReadyToSubmit;

namespace Topo.Services
{
    // What was read from Terrain for a unit; the report is worked out from it as often as the options change
    public class ReadyToSubmitData
    {
        public List<ReadyMemberInput> Members { get; set; } = new List<ReadyMemberInput>();
        public ReadyTemplates Templates { get; set; } = new ReadyTemplates();
    }

    public interface IReadyToSubmitService
    {
        public Task<ReadyToSubmitData> LoadAsync(List<MemberListModel> youth, string section, Action<string>? progress = null);
    }

    public class ReadyToSubmitService : IReadyToSubmitService
    {
        private static readonly string[] Types =
        {
            ReadyToSubmitBuilder.Type_Milestone, ReadyToSubmitBuilder.Type_Oas, ReadyToSubmitBuilder.Type_Sia,
            "adventurous_journey", "course_reflection", "personal_reflection"
        };

        private readonly ITerrainAPIService _terrainAPIService;
        // What a template asks does not change from member to member, so it is read once
        private readonly Dictionary<string, List<FormQuestion>> _formTemplates = new Dictionary<string, List<FormQuestion>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<FormQuestion>> _oasTemplates = new Dictionary<string, List<FormQuestion>>(StringComparer.OrdinalIgnoreCase);

        public ReadyToSubmitService(ITerrainAPIService terrainAPIService)
        {
            _terrainAPIService = terrainAPIService;
        }

        public async Task<ReadyToSubmitData> LoadAsync(List<MemberListModel> youth, string section, Action<string>? progress = null)
        {
            var data = new ReadyToSubmitData();
            var ordered = ReadyToSubmitBuilder.InPatrolOrder(youth);
            // the Reading line names each Scout the short way, as the report does
            var named = ordered.Select(m => new ReadyMemberInput { MemberId = m.id, MemberName = $"{m.first_name} {m.last_name}".Trim(), FirstName = m.first_name, LastName = m.last_name }).ToList();
            var shortNames = ReadyToSubmitBuilder.DisplayNames(named, false);
            for (var i = 0; i < ordered.Count; i++)
            {
                var member = ordered[i];
                progress?.Invoke($"{shortNames[named[i]]} ({i + 1} of {ordered.Count})");
                var responses = await Task.WhenAll(Types.Select(type => _terrainAPIService.GetAchievementsJsonAsync(member.id, type)));
                data.Members.Add(new ReadyMemberInput
                {
                    MemberId = member.id,
                    MemberName = $"{member.first_name} {member.last_name}".Trim(),
                    FirstName = member.first_name ?? "",
                    LastName = member.last_name ?? "",
                    Patrol = member.patrol_name ?? "",
                    Records = responses.SelectMany(ReadyParsers.ParseAchievements).ToList()
                });
            }

            // only the templates of what could be waiting are read
            var records = data.Members.SelectMany(m => m.Records).ToList();
            var formPaths = records.Where(r => ReadyToSubmitBuilder.NeedsFormTemplate(r, section)).Select(ReadyToSubmitBuilder.FormTemplatePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var oasPaths = records.Where(r => ReadyToSubmitBuilder.NeedsOasTemplate(r, section)).Select(r => r.Template).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var done = 0;
            foreach (var path in formPaths)
            {
                progress?.Invoke($"Checking what the forms ask ({++done} of {formPaths.Count + oasPaths.Count})");
                if (!_formTemplates.ContainsKey(path))
                {
                    var questions = ReadyParsers.ParseFormQuestions(await _terrainAPIService.GetTemplateJsonAsync(path));
                    // a template that couldn't be read is tried again next time
                    if (questions.Count > 0)
                        _formTemplates[path] = questions;
                }
                if (_formTemplates.TryGetValue(path, out var known))
                    data.Templates.Forms[path] = known;
            }
            foreach (var path in oasPaths)
            {
                progress?.Invoke($"Checking what the stages ask ({++done} of {formPaths.Count + oasPaths.Count})");
                if (!_oasTemplates.ContainsKey(path))
                {
                    var inputs = ReadyParsers.ParseOasInputs(await _terrainAPIService.GetTemplateJsonAsync(path));
                    if (inputs.Count > 0)
                        _oasTemplates[path] = inputs;
                }
                if (_oasTemplates.TryGetValue(path, out var known))
                    data.Templates.OasInputs[path] = known;
            }
            return data;
        }
    }
}
