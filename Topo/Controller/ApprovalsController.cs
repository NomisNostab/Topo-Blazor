using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Syncfusion.Blazor.Calendars;
using Syncfusion.Blazor.Grids;
using System.Net.NetworkInformation;
using System.Reflection;
using Topo.Model.AdditionalAwards;
using Topo.Model.Approvals;
using Topo.Model.ReportGeneration;
using Topo.Services;


namespace Topo.Controller
{
    public class ApprovalsController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public IApprovalsService _approvalsService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public ApprovalsPageViewModel model = new ApprovalsPageViewModel();

        public IEditorSettings DateEditParams = new DateEditCellParams
        {
            Params = new DatePickerModel() { ShowClearButton = true }
        };

        public SfGrid<ApprovalsListModel> GridInstance { get; set; }
        protected override void OnInitialized()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units;
            model.ApprovalSearchFromDate = DateTime.Now.AddMonths(-2);
            model.ApprovalSearchToDate = DateTime.Now;
            if (_storageService.UnitId != null)
            {
                model.UnitId = _storageService.UnitId;
                model.UnitName = _storageService.UnitName;
            }
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            model.UnitId = unitId;
            _storageService.UnitId = model.UnitId;
            model.UnitName = _storageService.UnitName;
            await RefreshApprovalsClick();
        }

        internal async Task RefreshApprovalsClick()
        {
            if (!string.IsNullOrEmpty(model.UnitName))
            {
                model.Approvals = await _approvalsService.GetApprovalListItems(_storageService.UnitId);
                model.Approvals = model.Approvals?.Where(a => a.submission_date >= model.ApprovalSearchFromDate && a.submission_date <= model.ApprovalSearchToDate.AddDays(1)).ToList() ?? new List<ApprovalsListModel>();
                if (model.ToBePresented)
                    model.Approvals = model.Approvals.Where(a => !string.IsNullOrEmpty(a.submission_outcome) && !a.presented_date.HasValue).ToList();
                if (model.IsPresented)
                    model.Approvals = model.Approvals.Where(a => a.presented_date.HasValue && a.presented_date != a.awarded_date).ToList();
                if (!model.ShowRejected)
                    model.Approvals = model.Approvals.Where(a => a.submission_outcome.ToLower() != "rejected").ToList();
            }
        }

        public async void ActionBeginHandler(ActionEventArgs<ApprovalsListModel> Args)
        {
            if (Args.RequestType == Syncfusion.Blazor.Grids.Action.Save)
            {
                var data = (ApprovalsListModel)Args.Data;
                await _approvalsService.UpdateApproval(model.UnitId, data);
            }
        }

        internal async Task ApprovalsReportPdfClick()
        {
            byte[] report = await ApprovalsReport(OutputType.PDF);
            var fileName = $"Approvals_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task ApprovalsReportXlsxClick()
        {
            byte[] report = await ApprovalsReport(OutputType.Excel);
            var fileName = $"Approvals_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        internal async Task<byte[]> ApprovalsReport(OutputType outputType = OutputType.PDF)
        {
            var filterSetings = GridInstance.FilterSettings;
            string[] nameFilterColumnValues = new string[0];
            string nameFilterOperator = "";
            string[] statusFilterColumnValues = new string[0];
            string statusFilterOperator = "";
            string groupCols = "";
            if (filterSetings != null && filterSetings.Columns != null)
            {
                var nameFilterSettings = filterSetings.Columns.Where(x => x.Field == "member_display_name");
                if (nameFilterSettings != null)
                {
                    nameFilterColumnValues = nameFilterSettings.Select(x => x.Value.ToString() ?? "").ToArray();
                    nameFilterOperator = nameFilterSettings.FirstOrDefault()?.Operator.ToString() ?? "";
                }
                var statusFilterSettings = filterSetings.Columns.Where(x => x.Field == "submission_status");
                if (statusFilterSettings != null)
                {
                    statusFilterColumnValues = statusFilterSettings.Select(x => x.Value.ToString() ?? "").ToArray();
                    statusFilterOperator = statusFilterSettings.FirstOrDefault()?.Operator.ToString() ?? "";
                }
            }
            var groupSettings = GridInstance.GroupSettings;
            if (groupSettings != null && groupSettings.Columns != null)
            {
                groupCols = string.Join(",", groupSettings.Columns);
            }

            var selectedApprovals = new List<ApprovalsListModel>();
            if (nameFilterOperator.ToLower() == "equal")
                selectedApprovals = model.Approvals.Where(t2 => nameFilterColumnValues.Count(m => m == t2.member_display_name) != 0).ToList();
            else
                selectedApprovals = model.Approvals.Where(t2 => nameFilterColumnValues.Count(m => m == t2.member_display_name) == 0).ToList();
            if (statusFilterOperator.ToLower() == "equal")
                selectedApprovals = selectedApprovals.Where(t2 => statusFilterColumnValues.Count(m => m == t2.submission_status) != 0).ToList();
            else
                selectedApprovals = selectedApprovals.Where(t2 => statusFilterColumnValues.Count(m => m == t2.submission_status) == 0).ToList();

            var groupByMember = (groupCols ?? "achievement_name") == "member_display_name";
            var groupName = _storageService.GroupName ?? "";
            var unitName = _storageService.UnitName ?? "";
            var section = _storageService.Section;

            var serialisedReportData = JsonConvert.SerializeObject(selectedApprovals);

            var report = await _reportService.GetApprovalsReport(groupName, section, unitName, outputType, serialisedReportData, model.ApprovalSearchFromDate, model.ApprovalSearchToDate, groupByMember);
            return report;
        }
    }
}
