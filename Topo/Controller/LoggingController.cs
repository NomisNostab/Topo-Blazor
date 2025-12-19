using Topo.Model.Index;
using Topo.Services;
using Microsoft.AspNetCore.Components;
using System.Reflection;
using Topo.Model.Loging;
using Topo.Model.ReportGeneration;
using System.Text;
using Microsoft.JSInterop;

namespace Topo.Controller
{
    public class LoggingController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; }

        [Inject]
        public LogQueueService _logQueue {  get; set; }

        [Inject]
        IJSRuntime JS { get; set; }

        public LoggingPageViewModel loggingPageViewModel = new LoggingPageViewModel();

        protected override void OnInitialized()
        {
            loggingPageViewModel.StartLogging = _logQueue.Queue.IsEmpty;
            loggingPageViewModel.DownloadLog = !_storageService.LogToQueue && !_logQueue.Queue.IsEmpty;
        }

        internal async Task StartLoggingClick()
        {
            _storageService.LogToQueue = true;
            loggingPageViewModel.StartLogging = false;
            loggingPageViewModel.DownloadLog = false;
        }

        internal async Task StopLoggingClick()
        {
            _storageService.LogToQueue = false;
            loggingPageViewModel.StartLogging = _logQueue.Queue.IsEmpty;
            loggingPageViewModel.DownloadLog = !_logQueue.Queue.IsEmpty;
        }

        internal async Task DownloadLogClick()
        {
            // Dump Log entries
            var queueDump = string.Join(Environment.NewLine, _logQueue.Queue);
            // Empty log queue
            _logQueue.Queue.Clear();
            // Send the data to JS to actually download the file
            await JS.InvokeVoidAsync("BlazorDownloadFile", $"{_storageService.GroupName}_{_storageService.UnitName}_Dump_{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.txt", "text/plain", Encoding.ASCII.GetBytes(queueDump));

            loggingPageViewModel.StartLogging = _logQueue.Queue.IsEmpty;
            loggingPageViewModel.DownloadLog = !_logQueue.Queue.IsEmpty;
        }
    }
}
