using Topo;
using Topo.Services;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.AspNetCore.Components;
using System.Globalization;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddSingleton<StorageService>();
builder.Services.AddScoped<ITerrainAPIService, TerrainAPIService>();
builder.Services.AddScoped<IMembersService, MembersService>();
builder.Services.AddScoped<ILoginService, LoginService>();
builder.Services.AddScoped<IReportService, ReportService>();
builder.Services.AddScoped<IProgramService, ProgramService>();
builder.Services.AddScoped<IOASService, OASService>();
builder.Services.AddScoped<ISIAService, SIAService>();
builder.Services.AddScoped<IMilestoneService, MilestoneService>();

builder.Services.AddScoped<SpinnerService>();
builder.Services.AddScoped<DisplaySpinnerAutomaticallyHttpMessageHandler>();
builder.Services.AddScoped(s =>
{
    var accessTokenHandler = s.GetRequiredService<DisplaySpinnerAutomaticallyHttpMessageHandler>();
    accessTokenHandler.InnerHandler = new HttpClientHandler();
    var uriHelper = s.GetRequiredService<NavigationManager>();
    return new HttpClient(accessTokenHandler)
    {
        BaseAddress = new Uri(uriHelper.BaseUri)
    };
});

System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("en-AU");
CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
CultureInfo.DefaultThreadCurrentUICulture = cultureInfo;


await builder.Build().RunAsync();
