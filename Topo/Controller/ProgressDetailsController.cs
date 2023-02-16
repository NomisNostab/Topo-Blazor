using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Topo.Model.Members;
using Topo.Model.Progress;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class ProgressDetailsController : ComponentBase
    {
        [Inject]
        public IProgressService _progressService { get; set; }

        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public IReportService _reportService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        [Parameter]
        public string MemberId { get; set; }

        public ProgressDetailsPageViewModel model = new ProgressDetailsPageViewModel();

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model = await _progressService.GetProgressDetailsPageViewModel(MemberId);
            model.GroupName = _storageService.GroupNameDisplay;
        }

        internal async Task ProgressPdfClick()
        {
            byte[] report = await ProgressReport(OutputType.PDF);
            var fileName = $"Personal_Progress_{model.Member.first_name}_{model.Member.last_name}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task ProgressXlsxClick()
        {
            byte[] report = await ProgressReport(OutputType.Excel);
            var fileName = $"Personal_Progress_{model.Member.first_name}_{model.Member.last_name}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> ProgressReport(OutputType outputType = OutputType.PDF)
        {
            var progressItems = model;

            var groupName = _storageService.GroupName ?? "";
            var unitName = _storageService.UnitName ?? "";
            var section = _storageService.Section;

            var serialisedProgressItems = JsonConvert.SerializeObject(progressItems);

            var report = await _reportService.GetProgressReport(groupName, section, unitName, outputType, serialisedProgressItems);
            return report;
        }
    }
}
