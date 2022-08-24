using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using System;
using System.Reflection.Metadata;
using System.Text.Json;
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
        IReportService _reportService { get; set; }

        public MembersPageViewModel model = new MembersPageViewModel();

        protected override async Task OnInitializedAsync()
        {
            model.Units = _storageService.Units;
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            model.UnitId = unitId; 
            _storageService.UnitId = model.UnitId;
            if (_storageService.Units != null)
                _storageService.UnitName = _storageService.Units.Where(u => u.Key == model.UnitId).FirstOrDefault().Value;
            var allMembers = await _membersService.GetMembersAsync(model.UnitId);
            model.Members = allMembers.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
            model.UnitName = _storageService.UnitName;
        }

        internal async Task PatrolListPdfClick()
        {

        }

        internal async Task PatrolListXlsxClick()
        {

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
        internal async Task PatrolSheetPdfClick()
        {

        }

        internal async Task PatrolSheetXlsxClick()
        {

        }

        private async Task<byte[]> MemberList(OutputType outputType = OutputType.PDF)
        {
            var groupName = _storageService.GroupName ?? "Group Name";
            var unitName = _storageService.UnitName ?? "Unit Name";
            var section = _storageService.Section;
            var sortedMemberList = model.Members.Where(m => m.isAdultLeader == 0).OrderBy(m => m.first_name).ThenBy(m => m.last_name).ToList();
            var serialisedSortedMemberList = JsonSerializer.Serialize(sortedMemberList);

            var report = await _reportService.GetMemberListReport(groupName, section, unitName, outputType, serialisedSortedMemberList);
            return report;
        }

    }
}
