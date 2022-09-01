using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using System.Reflection;
using Topo.Model.OAS;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class OASController : ComponentBase
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
        public IOASService _oasService { get; set; }

        public OASPageViewModel model = new OASPageViewModel();

        public ElementReference _select2Reference;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                await JS.InvokeVoidAsync("BindSelect2");
            }
        }

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.Units = _storageService.Units;

            var oasStageList = new List<OASStageListModel>();
            foreach (var stream in _oasService.GetOASStreamList())
            {
                var stageList = await _oasService.GetOASStageList(stream.Key);
                oasStageList.AddRange(stageList);
            }
            model.Stages = oasStageList;
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            model.UnitId = unitId;
            _storageService.UnitId = model.UnitId;
            if (_storageService.Units != null)
                _storageService.UnitName = _storageService.Units.Where(u => u.Key == model.UnitId).FirstOrDefault().Value;
            model.UnitName = _storageService.UnitName;
        }

        internal async Task OASWorksheetPdfClick()
        {
            if (!await GetSelections(_select2Reference))
                return;

            byte[] report = await OASWorksheet(OutputType.PDF);
            var fileName = $"OAS_Worksheet_{model.UnitName.Replace(' ', '_')}.pdf";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/pdf", report);
        }

        internal async Task OASWorksheetXlsxClick()
        {
            if (!await GetSelections(_select2Reference))
                return;

            byte[] report = await OASWorksheet(OutputType.Excel);
            var fileName = $"OAS_Worksheet_{model.UnitName.Replace(' ', '_')}.xlsx";

            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", fileName, "application/vnd.ms-excel", report);
        }

        private async Task<byte[]> OASWorksheet(OutputType outputType = OutputType.PDF)
        {
            var sortedAnswers = new List<OASWorksheetAnswers>();
            foreach (var selectedStageTemplate in model.SelectedStages)
            {
                var templateList = await _oasService.GetOASTemplate(selectedStageTemplate.Replace("/latest.json", ""));
                var selectedStage = model.Stages.Where(s => s.TemplateLink == selectedStageTemplate).FirstOrDefault() ?? new OASStageListModel();
                var sortedTemplateAnswers = await _oasService.GenerateOASWorksheetAnswers(model.UnitId, selectedStage, model.HideCompletedMembers, templateList);
                sortedAnswers.AddRange(sortedTemplateAnswers);
            }

            var groupName = _storageService.GroupName ?? "";
            var unitName = _storageService.UnitName ?? "";
            var section = _storageService.Section;

            var serialisedSortedMemberAnswers = JsonConvert.SerializeObject(sortedAnswers);

            var report = await _reportService.GetOASWorksheetReport(groupName, section, unitName, outputType, serialisedSortedMemberAnswers, model.BreakByPatrol);
            return report;
        }

        public async Task<bool> GetSelections(ElementReference elementReference)
        {
            model.SelectedStages = (await JS.InvokeAsync<List<string>>("getSelectedValues", _select2Reference)).ToArray<string>();
            model.StagesErrorMessage = "";
            if (model.SelectedStages == null || model.SelectedStages.Length == 0)
            {
                model.StagesErrorMessage = "Please select at least one stage";
                return false;
            }
            return true;
        }
    }

}
