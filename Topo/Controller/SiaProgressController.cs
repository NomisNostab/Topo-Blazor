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
        public bool Busy;
        public string Progress = "";

        public bool NeedsMember => model.Scope != SiaScope.WholeUnit;
        public bool NeedsProject => model.Scope == SiaScope.OneProject;

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
            if (string.IsNullOrEmpty(unitId))
            {
                model.Members = new List<Topo.Model.Members.MemberListModel>();
                return;
            }

            _storageService.UnitId = unitId;
            var allMembers = await _membersService.GetMembersAsync(unitId);
            model.Members = allMembers.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
        }

        internal async Task ScopeChange(SiaScope scope)
        {
            model.Scope = scope;
            model.ErrorMessage = "";
            if (scope == SiaScope.OneProject)
                await LoadProjects();
        }

        internal async Task MemberChange(ChangeEventArgs e)
        {
            model.MemberId = e.Value?.ToString() ?? "";
            model.ErrorMessage = "";
            await LoadProjects();
        }

        private async Task LoadProjects()
        {
            model.ProjectId = "";
            model.Projects = new List<SiaProjectChoice>();
            if (string.IsNullOrEmpty(model.MemberId) || model.Scope != SiaScope.OneProject)
                return;

            try
            {
                model.Projects = await _siaProgressService.GetProjectsAsync(model.MemberId, _storageService.Section);
                if (model.Projects.Count == 1)
                    model.ProjectId = model.Projects[0].Id;
                else if (model.Projects.Count == 0)
                    model.ErrorMessage = "This Scout has no Special Interest Area projects in this section.";
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't read the projects.";
            }
        }

        internal async Task PdfClick() => await Generate(OutputType.PDF);

        internal async Task XlsxClick() => await Generate(OutputType.Excel);

        private async Task Generate(OutputType outputType)
        {
            model.ErrorMessage = "";
            if (string.IsNullOrEmpty(model.UnitId))
                model.ErrorMessage = "Please select a unit.";
            else if (NeedsMember && string.IsNullOrEmpty(model.MemberId))
                model.ErrorMessage = "Please select a Scout.";
            else if (NeedsProject && string.IsNullOrEmpty(model.ProjectId))
                model.ErrorMessage = "Please select a project.";
            if (model.ErrorMessage != "")
                return;

            Busy = true;
            try
            {
                var request = new SiaProgressRequest
                {
                    Scope = model.Scope,
                    UnitName = _storageService.UnitName ?? "",
                    Section = _storageService.Section,
                    Members = model.Members,
                    MemberId = model.MemberId,
                    ProjectId = model.ProjectId,
                    ExcludeCompleted = model.ExcludeCompleted,
                    IncludeSummary = model.IncludeSummary,
                    DoubleSided = model.DoubleSided
                };
                var report = await _siaProgressService.GetReportAsync(request, step => { Progress = step; InvokeAsync(StateHasChanged); });
                if (report.Projects.Count == 0 && !report.IncludeSummary)
                {
                    model.ErrorMessage = model.ExcludeCompleted
                        ? "No Special Interest Area projects to show: everything found is completed."
                        : "No Special Interest Area projects to show.";
                    return;
                }

                var groupName = _storageService.GroupName ?? "Group Name";
                var bytes = await _reportService.GetSiaProgressReport(groupName, request.Section, request.UnitName, outputType, JsonConvert.SerializeObject(report));
                var isPdf = outputType == OutputType.PDF;
                var who = model.Scope == SiaScope.WholeUnit
                    ? request.UnitName
                    : model.Members.Where(m => m.id == model.MemberId).Select(m => $"{m.first_name} {m.last_name}").FirstOrDefault() ?? "";
                var fileName = $"SIA_Progress_{who.Trim().Replace(' ', '_')}" + (isPdf ? ".pdf" : ".xlsx");
                await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, isPdf ? "application/pdf" : "application/vnd.ms-excel", bytes);
            }
            catch (Exception)
            {
                model.ErrorMessage = "Couldn't build the report.";
            }
            finally
            {
                Busy = false;
                Progress = "";
            }
        }
    }
}
