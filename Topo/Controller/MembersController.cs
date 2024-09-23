using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.Login;
using Topo.Model.Members;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{

    public class MembersController : ComponentBase
    {
        [Inject]
        public IMembersService _membersService { get; set; }

        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public MembersPageViewModel model = new MembersPageViewModel();

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.Units = _storageService.Units;
            model.GroupName = _storageService.GroupNameDisplay;
            model.SuppressLastName = _storageService.SuppressLastName;
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
            if (string.IsNullOrEmpty(unitId))
            {
                model.UnitId = unitId;
                _storageService.UnitId = model.UnitId;
                model.UnitName = _storageService.UnitName;
                model.Members = new List<MemberListModel>();
                return;
            }
            model.UnitId = unitId;
            _storageService.UnitId = model.UnitId;
            var allMembers = await _membersService.GetMembersAsync(model.UnitId);
            model.Members = allMembers.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
            model.UnitName = _storageService.UnitName;
        }
        internal async Task PatrolListPdfClick()
        {
            byte[] report = await PatrolList(model.IncludeLeaders, OutputType.PDF);
            var fileName = $"Patrol_List_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task PatrolListXlsxClick()
        {
            byte[] report = await PatrolList(model.IncludeLeaders, OutputType.Excel);
            var fileName = $"Patrol_List_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> PatrolList(bool includeLeaders, OutputType outputType = OutputType.PDF)
        {
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = _storageService.UnitName ?? "Unit Name";
            var section = _storageService.Section;

            var allMembers = await _membersService.GetMembersAsync(model.UnitId);
            var sortedPatrolList = new List<MemberListModel>();
            if (includeLeaders)
                sortedPatrolList = allMembers.OrderBy(m => m.patrol_name).ToList();
            else
                sortedPatrolList = allMembers.Where(m => m.isAdultLeader == 0).OrderBy(m => m.patrol_name).ToList();
            var serialisedSortedMemberList = JsonConvert.SerializeObject(sortedPatrolList);

            var report = await _reportService.GetPatrolListReport(groupName, section, unitName, outputType, serialisedSortedMemberList, includeLeaders);
            return report;
        }

        internal async Task MemberListPdfClick()
        {
            byte[] report = await MemberList(OutputType.PDF);
            var fileName = $"Members_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task MemberListXlsxClick()
        {
            byte[] report = await MemberList(OutputType.Excel);
            var fileName = $"Members_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }
        private async Task<byte[]> MemberList(OutputType outputType = OutputType.PDF)
        {
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = _storageService.UnitName ?? "Unit Name";
            var section = _storageService.Section;
            var sortedMemberList = model.Members.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
            var serialisedSortedMemberList = JsonConvert.SerializeObject(sortedMemberList);

            var report = await _reportService.GetMemberListReport(groupName, section, unitName, outputType, serialisedSortedMemberList);
            return report;
        }

        internal async Task PatrolSheetPdfClick()
        {
            byte[] report = await PatrolSheet(OutputType.PDF);
            var fileName = $"Patrol_Sheets_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task PatrolSheetXlsxClick()
        {
            byte[] report = await PatrolSheet(OutputType.Excel);
            var fileName = $"Patrol_Sheets_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> PatrolSheet(OutputType outputType = OutputType.PDF)
        {
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = _storageService.UnitName ?? "Unit Name";
            var section = _storageService.Section;
            _storageService.SuppressLastName = model.SuppressLastName;
            List<MemberListModel> sortedMemberList = new List<MemberListModel>();
            foreach (var member in model.Members.Where(m => m.isAdultLeader == 0).OrderBy(m => m.patrol_name))
            {
                string lastName = _storageService.SuppressLastName ? member.last_name.Substring(0, 1).ToUpper() : member.last_name;
                MemberListModel memberCopy = new MemberListModel
                {
                    patrol_name = member.patrol_name,
                    first_name = member.first_name,
                    last_name = lastName,
                    isAdultLeader = member.isAdultLeader,
                    patrol_duty = member.patrol_duty
                };
                sortedMemberList.Add(memberCopy);
            }
            var serialisedSortedMemberList = JsonConvert.SerializeObject(sortedMemberList);

            var report = await _reportService.GetPatrolSheetsReport(groupName, section, unitName, outputType, serialisedSortedMemberList);
            return report;
        }

        internal async Task MemberRefreshClick()
        {
            _membersService.ClearMemberCache(_storageService.UnitId);
            await UnitChange(_storageService.UnitId);
        }
    }
}
