using Topo.Services;

namespace Topo
{
    public class DisplaySpinnerAutomaticallyHttpMessageHandler : DelegatingHandler
    {
        private readonly SpinnerService _spinnerService;
        public DisplaySpinnerAutomaticallyHttpMessageHandler(SpinnerService spinnerService)
        {
            _spinnerService = spinnerService;
        }
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            _spinnerService.Show();
            //  await Task.Delay(1000);
            var response = await base.SendAsync(request, cancellationToken);
            _spinnerService.Hide();
            return response;
        }
    }
}
