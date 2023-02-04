using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using Newtonsoft.Json;
using Syncfusion.Blazor.Calendars;
using Syncfusion.Blazor.Grids;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using Topo.Model.AdditionalAwards;
using Topo.Model.Approvals;
using Topo.Model.ReportGeneration;
using Topo.Services;

namespace Topo.Controller
{
    public class ApprovalsBackupController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public IApprovalsService _approvalsService { get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public BackupRestorePageViewModel model = new BackupRestorePageViewModel();

        protected override void OnInitialized()
        {
            if (!_storageService.IsAuthenticated)
                NavigationManager.NavigateTo("index");

            model.Units = _storageService.Units;
            if (_storageService.UnitId != null)
            {
                model.SelectedUnitId = _storageService.UnitId;
                if (_storageService.Units != null)
                    _storageService.UnitName = _storageService.Units.Where(u => u.Key == model.SelectedUnitId).FirstOrDefault().Value;
                model.SelectedUnitName = _storageService.UnitName;
            }
            model.GroupName = _storageService.GroupNameDisplay;
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            var unitId = e.Value?.ToString() ?? "";
            model.SelectedUnitId = unitId;
            _storageService.UnitId = model.SelectedUnitId;
            if (_storageService.Units != null)
                _storageService.UnitName = _storageService.Units.Where(u => u.Key == model.SelectedUnitId).FirstOrDefault().Value;
            model.SelectedUnitName = _storageService.UnitName;
        }

        internal async Task DownloadApprovalsClick()
        {
            if (!string.IsNullOrEmpty(model.SelectedUnitName))
            {
                var downloadJson = await _approvalsService.DownloadApprovalList(model.SelectedUnitId);
                // Send the data to JS to actually download the file
                await JS.InvokeVoidAsync("BlazorDownloadFile", $"Approvals_Backup_{model.SelectedUnitName.Replace(' ', '_')}_{DateTime.Now.ToString("yyyy-MM-dd")}.json", "application/json", Encoding.ASCII.GetBytes(downloadJson));
            }
        }

        internal async Task LoadFiles(InputFileChangeEventArgs e)
        {
            model.approvalsFile = e.File;
        }

        internal async Task UploadApprovalsClick()
        {
            if (!string.IsNullOrEmpty(model.SelectedUnitName) && model.approvalsFile != null && model.approvalsFile.Size > 0)
            {
                _approvalsService.UploadApprovals(model.approvalsFile, model.SelectedUnitId);
                NavigationManager.NavigateTo("approvals/approvals");
            }
        }
    }
}
