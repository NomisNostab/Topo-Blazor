using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;
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

            model.GroupName = _storageService.GroupNameDisplay;
            model.Units = _storageService.Units;
            model.Stages = await _oasService.GetOASStagesList();
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
            _storageService.UnitId = model.UnitId;
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

            var report = await _reportService.GetOASWorksheetReport(groupName, section, unitName, outputType, serialisedSortedMemberAnswers, model.BreakByPatrol, model.FormatLikeTerrain);
            return report;
        }

        internal async Task GetCoreSelections(ChangeEventArgs arg)
        {
            model.SelectedStages = new string[0];
            model.StagesErrorMessage = "";
            if (model.UseCore)
            {
                List<OASStageListModel> coreStages = new List<OASStageListModel>();
                for (int i = 0; i < 9; i++)
                {
                    if (model.CoreStages[i])
                    {
                        coreStages = coreStages.Concat(GetCoreForStage(i + 1)).ToList();
                    }
                }

                foreach (var coreStage in coreStages)
                {
                    model.SelectedStages = model.SelectedStages.Append(coreStage.TemplateLink).ToArray();
                }

                if (model.SelectedStages == null || model.SelectedStages.Length == 0)
                {
                    model.StagesErrorMessage = "Please select at least one stage";
                }
            }
            else
            {
                for (int i = 0; i < 9; i++)
                {
                    model.CoreStages[i] = false;
                }
                model.StagesErrorMessage = "Please select at least one stage";
            }

            await JS.InvokeVoidAsync("BindSelect2");
            await JS.InvokeVoidAsync("BindSelect2");
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

        private List<OASStageListModel> GetCoreForStage(int stageNumber)
        {
            var stages = model.Stages.Where(stage => stage.Stage == stageNumber && (stage.Stream == "Bushcraft" || stage.Stream == "Bushwalking" || stage.Stream == "Camping")).ToList();
            return stages;
        }
    }

}
