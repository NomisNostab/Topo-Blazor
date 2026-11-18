using Blazored.LocalStorage;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using System.Text;
using Topo.Model.Members;
using Topo.Model.PatrolPlanner;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    // Backs the Patrol Planner drag-and-drop board. Everything here is a local, in-memory "what if"
    // plan: there is no Terrain API to write patrol membership back, so nothing is saved to Terrain
    // automatically. The leader applies the plan by hand, using the exported patrol list/changes CSV.
    public class PatrolPlannerController : ComponentBase
    {
        [Inject]
        public IMembersService _membersService { get; set; }

        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public ILocalStorageService _localStorageService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public PatrolPlannerPageViewModel model = new PatrolPlannerPageViewModel();

        private string? _draggedMemberId;
        private List<string> _originalPatrolNames = new();
        private Dictionary<string, string> _patrolColorOverrides = new();

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.Units = _storageService.Units;
            model.GroupName = _storageService.GroupNameDisplay;
            if (!string.IsNullOrEmpty(_storageService.UnitId))
            {
                await UnitChange(_storageService.UnitId);
            }
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            await UnitChange(unitId);
        }

        internal async Task UnitChange(string unitId)
        {
            model.UnitId = unitId;
            _storageService.UnitId = unitId;
            model.UnitName = _storageService.UnitName;
            if (string.IsNullOrEmpty(unitId))
            {
                model.Members = new List<PatrolPlannerMember>();
                model.PatrolNames = new List<string>();
                _originalPatrolNames = new List<string>();
                return;
            }

            _patrolColorOverrides = await _localStorageService.GetItemAsync<Dictionary<string, string>>($"{unitId}_PatrolColors") ?? new Dictionary<string, string>();

            var allMembers = await _membersService.GetMembersAsync(unitId);
            model.Members = allMembers
                .Where(m => m.isAdultLeader == 0)
                .Select(m => new PatrolPlannerMember
                {
                    Id = m.id,
                    FirstName = m.first_name,
                    LastName = m.last_name,
                    Age = m.age,
                    UnitCouncil = m.unit_council,
                    OriginalUnitCouncil = m.unit_council,
                    Gender = m.gender,
                    OriginalPatrolName = m.patrol_name,
                    OriginalPatrolDuty = m.patrol_duty,
                    PatrolName = m.patrol_name,
                    PatrolDuty = m.patrol_duty
                })
                .ToList();

            model.PatrolNames = model.Members
                .Select(m => m.OriginalPatrolName)
                .Where(n => !string.IsNullOrEmpty(n))
                .Distinct()
                .OrderBy(n => n)
                .ToList();
            _originalPatrolNames = new List<string>(model.PatrolNames);
        }

        internal async Task RefreshClick()
        {
            _membersService.ClearMemberCache(model.UnitId);
            await UnitChange(model.UnitId);
        }

        // Undoes member moves/role/council changes AND any patrols added or removed this session,
        // by restoring model.PatrolNames from the snapshot UnitChange took when it first loaded.
        internal void ResetChanges()
        {
            foreach (var member in model.Members)
            {
                member.PatrolName = member.OriginalPatrolName;
                member.PatrolDuty = member.OriginalPatrolDuty;
                member.UnitCouncil = member.OriginalUnitCouncil;
            }
            model.PatrolNames = new List<string>(_originalPatrolNames);
        }

        internal void AddPatrolClick()
        {
            var name = model.NewPatrolName?.Trim() ?? "";
            if (string.IsNullOrEmpty(name))
                return;
            if (!model.PatrolNames.Any(p => string.Equals(p, name, StringComparison.OrdinalIgnoreCase)))
            {
                model.PatrolNames.Add(name);
            }
            model.NewPatrolName = "";
        }

        internal void RemovePatrol(string patrolName)
        {
            model.PatrolNames.Remove(patrolName);
            foreach (var member in model.Members.Where(m => m.PatrolName == patrolName))
            {
                member.PatrolName = "";
            }
            // Any saved colour override for this patrol name is deliberately left in local storage:
            // harmless if the patrol is gone for good, and remembered for free if it's re-added later.
        }

        // A Unit Leader still has a patrol_name in Terrain's data (the UL role sits above patrols,
        // not inside them), so they're pulled out into their own banner here rather than appearing
        // in a column. GetColumnMembers excludes them for the same reason.
        internal IEnumerable<PatrolPlannerMember> ULMembers =>
            model.Members.Where(m => m.PatrolDuty == "UL").OrderBy(m => m.FirstName).ThenBy(m => m.LastName);

        internal IEnumerable<string> AllColumnNames
        {
            get
            {
                var columns = new List<string>(model.PatrolNames);
                if (model.Members.Any(m => string.IsNullOrEmpty(m.PatrolName)))
                {
                    columns.Add(string.Empty);
                }
                return columns;
            }
        }

        // PL first, then APL, then everyone else oldest-to-youngest - matches how a patrol is
        // normally listed on paper (leaders at top, members by seniority underneath).
        internal IEnumerable<PatrolPlannerMember> GetColumnMembers(string patrolName)
        {
            return model.Members
                .Where(m => m.PatrolName == patrolName && m.PatrolDuty != "UL")
                .OrderBy(m => m.PatrolDuty == "PL" ? 0 : m.PatrolDuty == "APL" ? 1 : 2)
                .ThenByDescending(m => GetAgeInMonths(m.Age));
        }

        internal IEnumerable<PatrolPlannerMember> ChangedMembers =>
            model.Members.Where(m => m.IsChanged).OrderBy(m => m.FirstName).ThenBy(m => m.LastName);

        internal bool HasChanges => model.Members.Any(m => m.IsChanged);

        private static int GetAgeInMonths(string age)
        {
            // age is formatted like "14y 9m"
            var parts = age.Replace("y", "").Replace("m", "").Split(' ');
            if (parts.Length == 2 && int.TryParse(parts[0], out var years) && int.TryParse(parts[1], out var months))
            {
                return years * 12 + months;
            }
            return 0;
        }

        private static readonly string[] PatrolPalette = new[]
        {
            "#CFE3FF", "#FFE0C2", "#D4F0D0", "#F6D6E0", "#E3D6FF",
            "#FFF3B0", "#C9EFEF", "#E6D3B3", "#D9D9F3", "#FFD9D9"
        };

        // Light tints to use when a patrol is literally named after a colour (e.g. "Red Patrol").
        private static readonly Dictionary<string, string> NamedPatrolColors = new(StringComparer.OrdinalIgnoreCase)
        {
            { "red", "#FFD9D9" },
            { "blue", "#CFE3FF" },
            { "green", "#D4F0D0" },
            { "yellow", "#FFF3B0" },
            { "orange", "#FFE0C2" },
            { "purple", "#E3D6FF" },
            { "violet", "#E3D6FF" },
            { "indigo", "#D9D9F3" },
            { "pink", "#F6D6E0" },
            { "magenta", "#F6D6E0" },
            { "brown", "#E6D3B3" },
            { "tan", "#E6D3B3" },
            { "cyan", "#C9EFEF" },
            { "teal", "#C9EFEF" },
            { "turquoise", "#C9EFEF" },
            { "gold", "#FFE9A8" },
            { "maroon", "#E8C2C2" },
            { "navy", "#C2D4E8" },
            { "black", "#D9D9D9" },
            { "white", "#F5F5F5" },
            { "grey", "#E2E2E2" },
            { "gray", "#E2E2E2" },
            { "silver", "#E8E8E8" }
        };

        // Colour priority: a colour the leader picked for this patrol (saved locally) wins first,
        // then a colour word found in the patrol's own name (e.g. "Drop Bear (Green)"), and only
        // then a stable colour assigned purely by the patrol's position in the alphabetised list.
        internal string GetPatrolColor(string patrolName)
        {
            if (string.IsNullOrEmpty(patrolName))
                return "#E2E2E2";

            if (_patrolColorOverrides.TryGetValue(patrolName, out var overrideColor) && !string.IsNullOrEmpty(overrideColor))
                return overrideColor;

            // Split on brackets/punctuation too, since real patrol names often read "Name (Colour)".
            var words = patrolName.Split(new[] { ' ', '-', '_', '(', ')', ',', '.' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var word in words)
            {
                if (NamedPatrolColors.TryGetValue(word, out var namedColor))
                    return namedColor;
            }

            var index = model.PatrolNames.IndexOf(patrolName);
            if (index < 0)
                return "#E2E2E2";
            return PatrolPalette[index % PatrolPalette.Length];
        }

        // Patrol colour choices are a local display preference, not Terrain data, so they're kept in
        // this browser's local storage (keyed by unit) rather than anywhere near the member/report models.
        internal async Task SetPatrolColor(string patrolName, string colorHex)
        {
            if (string.IsNullOrEmpty(patrolName) || string.IsNullOrEmpty(colorHex))
                return;
            _patrolColorOverrides[patrolName] = colorHex;
            await _localStorageService.SetItemAsync($"{model.UnitId}_PatrolColors", _patrolColorOverrides);
        }

        // Counts use the same M/F/NB labels as the card badges (GetColumnMembers/PatrolMemberCard)
        // so the tally and the cards never disagree on how gender is written.
        internal (int total, int male, int female, int nonBinary, int unspecified, string averageAge) GetColumnStats(string patrolName)
        {
            var members = GetColumnMembers(patrolName).ToList();
            var averageAge = "";
            if (members.Any())
            {
                var avgMonths = (int)Math.Round(members.Average(m => GetAgeInMonths(m.Age)));
                averageAge = $"{avgMonths / 12}y {avgMonths % 12}m";
            }
            return (members.Count,
                members.Count(m => m.Gender == "M"),
                members.Count(m => m.Gender == "F"),
                members.Count(m => m.Gender == "NB"),
                members.Count(m => string.IsNullOrEmpty(m.Gender)),
                averageAge);
        }

        // Flags a column's count badge when it's more than one member away from the average across
        // real patrols (Unassigned is excluded from the average - it isn't meant to be "balanced").
        internal string GetCountClass(string patrolName)
        {
            var countsByColumn = AllColumnNames
                .Where(c => !string.IsNullOrEmpty(c))
                .Select(c => GetColumnMembers(c).Count())
                .ToList();
            if (!countsByColumn.Any() || string.IsNullOrEmpty(patrolName))
                return string.Empty;

            var target = (double)countsByColumn.Sum() / countsByColumn.Count;
            var count = GetColumnMembers(patrolName).Count();
            if (Math.Abs(count - target) > 1)
                return "patrol-count-unbalanced";
            return string.Empty;
        }

        internal void OnDragStart(string memberId)
        {
            _draggedMemberId = memberId;
        }

        internal void OnDropToPatrol(string patrolName)
        {
            if (string.IsNullOrEmpty(_draggedMemberId))
                return;

            var member = model.Members.FirstOrDefault(m => m.Id == _draggedMemberId);
            if (member != null)
            {
                member.PatrolName = patrolName;
            }
            _draggedMemberId = null;
        }

        internal void OnDutyChange(PatrolPlannerMember member, string duty)
        {
            member.PatrolDuty = duty;
            // Promoting someone adds them to Unit Council automatically, per the group's convention.
            // Demoting them back to Member does NOT remove them again - that's a separate, deliberate
            // decision a leader should make on its own, not an automatic side effect of this control.
            if (duty == "APL" || duty == "PL" || duty == "UL")
            {
                member.UnitCouncil = true;
            }
        }

        // Cycles unset -> M -> F -> NB -> unset. "NB" (Non-Binary) is a real, explicit state here,
        // not lumped in with "unset" - a leader can tell the two apart in both the tally and the export.
        internal async Task CycleGenderClick(PatrolPlannerMember member)
        {
            member.Gender = member.Gender switch
            {
                "M" => "F",
                "F" => "NB",
                "NB" => "",
                _ => "M"
            };
            await _membersService.SetMemberGenderAsync(model.UnitId, member.Id, member.Gender);
        }

        internal async Task PatrolListPdfClick()
        {
            byte[] report = await GenerateUpdatedPatrolList(OutputType.PDF);
            var fileName = $"Patrol_List_{model.UnitName.Replace(' ', '_')}.pdf";
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task PatrolListXlsxClick()
        {
            byte[] report = await GenerateUpdatedPatrolList(OutputType.Excel);
            var fileName = $"Patrol_List_{model.UnitName.Replace(' ', '_')}.xlsx";
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        // Reuses the existing PatrolList report type (same one Members.razor already calls) rather than
        // adding a new ReportType: that pipeline is served by the TopoReportFunction Lambda, which would
        // need a separate deployment to pick up a brand new report type. This just feeds it the reshuffled
        // in-memory list instead of Terrain's live data, so it works today with no extra deployment step.
        private async Task<byte[]> GenerateUpdatedPatrolList(OutputType outputType)
        {
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = model.UnitName;
            var section = _storageService.Section;

            var updatedList = model.Members
                .Select(m => new MemberListModel
                {
                    first_name = m.FirstName,
                    last_name = m.LastName,
                    patrol_name = m.PatrolName,
                    patrol_duty = m.PatrolDuty,
                    unit_council = m.UnitCouncil,
                    isAdultLeader = 0
                })
                .OrderBy(m => m.patrol_name)
                .ToList();
            var serialisedList = JsonConvert.SerializeObject(updatedList);

            return await _reportService.GetPatrolListReport(groupName, section, unitName, outputType, serialisedList, false);
        }

        // Built entirely client-side (plain CSV text) rather than through the report Lambda, since it's
        // just a flat list of changed rows and doesn't need Syncfusion's PDF/XLSX layout engine.
        internal async Task DownloadChangesClick()
        {
            var changed = ChangedMembers.ToList();
            var sb = new StringBuilder();
            sb.AppendLine("First Name,Last Name,Duty,Patrol,Unit Council");
            foreach (var member in changed)
            {
                sb.AppendLine(string.Join(",",
                    CsvField(member.FirstName),
                    CsvField(member.LastName),
                    CsvField(member.PatrolDuty),
                    CsvField(DescribePatrolChange(member)),
                    CsvField(member.UnitCouncil ? "Yes" : "")));
            }

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            var fileName = $"Patrol_Changes_{model.UnitName.Replace(' ', '_')}.csv";
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "text/csv", bytes);
        }

        // Used for the "Changes to make in Terrain" list/CSV only (not the board itself), where the
        // text needs to read as an instruction. A member who had a patrol and now has none needs an
        // active removal in Terrain; one who was never assigned a patrol needs no action there at all.
        internal static string DescribePatrolChange(PatrolPlannerMember member)
        {
            if (string.IsNullOrEmpty(member.PatrolName))
            {
                return string.IsNullOrEmpty(member.OriginalPatrolName)
                    ? "Unassigned"
                    : "Remove from patrol";
            }
            return member.PatrolName;
        }

        private static string CsvField(string value)
        {
            if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            {
                return $"\"{value.Replace("\"", "\"\"")}\"";
            }
            return value;
        }
    }
}
