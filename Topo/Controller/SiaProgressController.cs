using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.ReportGeneration;
using Topo.Model.SiaProgress;
using Topo.Services;

namespace Topo.Controller
{
    public class SiaProgressController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        [Inject]
        public IMembersService _membersService { get; set; }

        [Inject]
        public ISiaProgressService _siaProgressService { get; set; }

        public SiaProgressPageViewModel model = new SiaProgressPageViewModel();
        public bool Loading;
        public string Progress = "";
        // Worked out again whenever an option changes, from what was read for the Scout or unit
        public SiaProgressReport? Report;
        private SiaProgressInput? _data;

        public bool NeedsMember => model.Scope != SiaScope.WholeUnit;
        public bool NeedsProject => model.Scope == SiaScope.OneProject;
        // Nothing to download when no project was found, unless the summary was asked for
        public bool CanExport => Report != null && (Report.Projects.Count > 0 || Report.IncludeSummary);

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
            {
                NavigationManager.NavigateTo("index");
                return;
            }

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units;
            if (!string.IsNullOrEmpty(_storageService.UnitId))
                await UnitChange(_storageService.UnitId);
        }

        internal async Task UnitChange(ChangeEventArgs e) => await UnitChange(e.Value?.ToString() ?? "");

        internal async Task UnitChange(string unitId)
        {
            model.UnitId = unitId;
            model.MemberId = "";
            model.ProjectId = "";
            model.Projects = new List<SiaProjectChoice>();
            model.ErrorMessage = "";
            Clear();
            if (string.IsNullOrEmpty(unitId))
            {
                model.Members = new List<Topo.Model.Members.MemberListModel>();
                return;
            }

            _storageService.UnitId = unitId;
            var allMembers = await _membersService.GetMembersAsync(unitId);
            model.Members = allMembers.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
            await Load();
        }

        internal async Task ScopeChange(SiaScope scope)
        {
            model.Scope = scope;
            model.ErrorMessage = "";
            await Load();
        }

        internal async Task MemberChange(ChangeEventArgs e)
        {
            model.MemberId = e.Value?.ToString() ?? "";
            model.ErrorMessage = "";
            await Load();
        }

        private void Clear()
        {
            Report = null;
            _data = null;
        }

        // Reads the projects of the Scout, or of the whole unit, once; the preview then follows the options
        private async Task Load()
        {
            Clear();
            model.ProjectId = "";
            model.Projects = new List<SiaProjectChoice>();
            if (string.IsNullOrEmpty(model.UnitId) || (NeedsMember && string.IsNullOrEmpty(model.MemberId)))
                return;

            Loading = true;
            try
            {
                _data = await _siaProgressService.LoadAsync(new SiaProgressRequest
                {
                    Scope = model.Scope,
                    UnitName = _storageService.UnitName ?? "",
                    Section = _storageService.Section,
                    UnitId = model.UnitId,
                    Members = model.Members,
                    MemberId = model.MemberId
                }, step => { Progress = step; InvokeAsync(StateHasChanged); });

                if (NeedsProject)
                {
                    model.Projects = SiaProgressBuilder.ProjectChoices(_data.Members.FirstOrDefault()?.Results ?? new Topo.Model.SIA.GetSIAResultsModel(), _storageService.Section);
                    if (model.Projects.Count == 1)
                        model.ProjectId = model.Projects[0].Id;
                    else if (model.Projects.Count == 0)
                        model.ErrorMessage = "This Scout has no Special Interest Area projects in this section.";
                }
                Recalculate();
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't read the projects from Terrain.";
            }
            finally
            {
                Loading = false;
                Progress = "";
            }
        }

        // The preview follows the options as they are changed, without reading Terrain again
        internal void OptionChanged() => Recalculate();

        private void Recalculate()
        {
            Report = null;
            // one project needs choosing first, from the Scout's list
            if (_data == null || (NeedsProject && string.IsNullOrEmpty(model.ProjectId)))
                return;

            _data.ProjectId = NeedsProject ? model.ProjectId : "";
            _data.ExcludeCompleted = model.ExcludeCompleted;
            _data.IncludeSummary = model.IncludeSummary;
            _data.DoubleSided = model.DoubleSided;
            Report = SiaProgressBuilder.Build(_data);
        }

        internal async Task PdfClick() => await Generate(OutputType.PDF);

        internal async Task XlsxClick() => await Generate(OutputType.Excel);

        private async Task Generate(OutputType outputType)
        {
            if (!CanExport)
                return;

            try
            {
                var unitName = _storageService.UnitName ?? "";
                var groupName = _storageService.GroupName ?? "Group Name";
                var bytes = await _reportService.GetSiaProgressReport(groupName, _storageService.Section, unitName, outputType, JsonConvert.SerializeObject(Report));
                var isPdf = outputType == OutputType.PDF;
                var who = model.Scope == SiaScope.WholeUnit
                    ? unitName
                    : model.Members.Where(m => m.id == model.MemberId).Select(m => $"{m.first_name} {m.last_name}").FirstOrDefault() ?? "";
                var fileName = $"SIA_Progress_{who.Trim().Replace(' ', '_')}" + (isPdf ? ".pdf" : ".xlsx");
                await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, isPdf ? "application/pdf" : "application/vnd.ms-excel", bytes);
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't build the report.";
            }
        }
    }
}
