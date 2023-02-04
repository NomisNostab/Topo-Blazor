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

    }
}
