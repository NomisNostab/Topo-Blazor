using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using System;
using Topo.Model.Members;
using Topo.Model.OAS;
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
        public IOASService _oasService { get; set; }

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
            model.DisableOAS = model.OASSummaries.Where(o => o.Awarded == DateTime.MinValue).Count() == 0 ? "disable" : null;
            model.DisableCoreOAS = model.OASSummaries
                .Where(o => o.Stream == "bushcraft" || o.Stream == "bushwalking" || o.Stream == "camping")
                .Where(o => o.Awarded == DateTime.MinValue).Count() == 0 ? "disable" : null;
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

        internal async Task StartedOasPdfClick()
        {
            byte[] report = await StartedOASWorksheet(OutputType.PDF);
            var fileName = $"Personal_Progress_{model.Member.first_name}_{model.Member.last_name}_OAS.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task StartedOasXlsxClick()
        {
            byte[] report = await StartedOASWorksheet(OutputType.Excel);
            var fileName = $"Personal_Progress_{model.Member.first_name}_{model.Member.last_name}_OAS.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        internal async Task StartedCoreOasPdfClick()
        {
            byte[] report = await StartedCoreOASWorksheet(OutputType.PDF);
            var fileName = $"Personal_Progress_{model.Member.first_name}_{model.Member.last_name}_OAS.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }
        internal async Task StartedCoreOasXlsxClick()
        {
            byte[] report = await StartedCoreOASWorksheet(OutputType.Excel);
            var fileName = $"Personal_Progress_{model.Member.first_name}_{model.Member.last_name}_OAS.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> StartedOASWorksheet(OutputType outputType = OutputType.PDF)
        {
            var startedOASStages = model.OASSummaries.Where(o => o.Awarded == DateTime.MinValue || o.Approved == DateTime.MinValue).ToList();
            return await OASWorksheet(startedOASStages, outputType);
        }

        private async Task<byte[]> StartedCoreOASWorksheet(OutputType outputType = OutputType.PDF)
        {
            var startedOASStages = model.OASSummaries.Where(o => o.Awarded == DateTime.MinValue || o.Approved == DateTime.MinValue)
                                                .Where(o => o.Stream == "bushcraft" || o.Stream == "bushwalking" || o.Stream == "camping")
                                                .ToList();
            return await OASWorksheet(startedOASStages, outputType);
        }

        private async Task<byte[]> OASWorksheet(List<OASSummary> startedOASStages, OutputType outputType = OutputType.PDF)
        {
            var stages = await _oasService.GetOASStagesList();
            var sortedAnswers = new List<OASWorksheetAnswers>();
            foreach (var selectedStageTemplate in startedOASStages)
            {
                var templateList = await _oasService.GetOASTemplate(selectedStageTemplate.Template);
                var selectedStage = stages.Where(s => s.TemplateLink == $"{selectedStageTemplate.Template}/latest.json").FirstOrDefault() ?? new OASStageListModel();
                var sortedTemplateAnswers = await _oasService.GenerateOASWorksheetAnswersForMember(_storageService.UnitId, selectedStage, false, templateList, model.Member.member_number);
                sortedAnswers.AddRange(sortedTemplateAnswers);
            }

            var groupName = _storageService.GroupName ?? "";
            var unitName = _storageService.UnitName ?? "";
            var section = _storageService.Section;

            var serialisedSortedMemberAnswers = JsonConvert.SerializeObject(sortedAnswers);

            var report = await _reportService.GetOASWorksheetReport(groupName, section, unitName, outputType, serialisedSortedMemberAnswers, false, true);
            return report;
        }

        public string FormatOAS(string stream, int stage)
        {
            if (model.OASSummaries.Where(o => o.Stream == stream && o.Stage == stage).FirstOrDefault() == null)
            {
                return "";
            }
            else if (model.OASSummaries.Where(o => o.Stream == stream && o.Stage == stage && o.Awarded > DateTime.MinValue).OrderByDescending(o => o.Awarded).FirstOrDefault() != null)
            {
                var summary = model.OASSummaries.Where(o => o.Stream == stream && o.Stage == stage).OrderByDescending(o => o.Awarded).FirstOrDefault();
                return $"{summary.Awarded.ToString("dd/MM/yy")} {summary.Section}";
            }
            else if (model.OASSummaries.Where(o => o.Stream == stream && o.Stage == stage && o.Approved > DateTime.MinValue).OrderByDescending(o => o.Approved).FirstOrDefault() != null)
            {
                var summary = model.OASSummaries.Where(o => o.Stream == stream && o.Stage == stage).OrderByDescending(o => o.Approved).FirstOrDefault();
                return "Approved";
//return $"{summary.Awarded.ToString("dd/MM/yy")} {summary.Section}";
            }
            else
            {
                return "Started";
            }
        }
    }
}
