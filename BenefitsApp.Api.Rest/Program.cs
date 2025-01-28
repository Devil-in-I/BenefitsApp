using BenefitsApp.Core;
using BenefitsApp.Core.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddDbContext(builder.Configuration);
builder.Services.AddBenefitsCoreModule();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

var benefitsGroup = app.MapGroup("benefits");
benefitsGroup
    .MapGet("", async (IBenefitService benefitService) =>
    {
        return await benefitService.GetAllBenefitsAsync();
    })
    .WithName("Get all benefits")
    .WithOpenApi();

benefitsGroup
    .MapGet("/count", (IBenefitService benefitsService) =>
    {
        return benefitsService.GetBenefitsCount();
    })
    .WithName("Get benefits count")
    .WithOpenApi();

app.Run();