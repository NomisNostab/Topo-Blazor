using Topo;
using Topo.Services;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddSingleton<StorageService>();
builder.Services.AddScoped<ITerrainAPIService, TerrainAPIService>();
builder.Services.AddScoped<IMembersService, MembersService>();
builder.Services.AddScoped<ILoginService, LoginService>();
builder.Services.AddScoped<IReportService, ReportService>();

await builder.Build().RunAsync();
