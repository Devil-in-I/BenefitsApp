using BenefitsApp.Core.Clients;
using BenefitsApp.Core.Models;
using BenefitsApp.UI.Components;
using PnP.Core.Auth;
using PnP.Core.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.Configure<SharePointCredentialsOptions>(
    builder.Configuration.GetSection(SharePointCredentialsOptions.SharePointCredentials));

builder.Services.AddScoped<ISharePointService, SharePointService>();

builder.Services.AddPnPCore(options => options.DefaultAuthenticationProvider = new InteractiveAuthenticationProvider());

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
