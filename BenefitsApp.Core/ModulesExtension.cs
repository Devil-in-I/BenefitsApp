using BenefitsApp.Core.Configuration;
using BenefitsApp.Core.Data;
using BenefitsApp.Core.Repositories;
using BenefitsApp.Core.Repositories.Interfaces;
using BenefitsApp.Core.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace BenefitsApp.Core
{
    public static class ModulesExtension
    {
        public static IServiceCollection AddDbContext(this IServiceCollection services, IConfiguration configuration)
        {
            var mongoDBSettings = configuration.GetRequiredSection(nameof(MongoDbConfiguration)).Get<MongoDbConfiguration>()
                ?? throw new NullReferenceException();

            services.AddDbContext<BenefitsDbContext>(options =>
            options.UseMongoDB(mongoDBSettings.AtlasURI ?? "", mongoDBSettings.DatabaseName ?? ""));

            return services;
        }

        public static IServiceCollection AddBenefitsCoreModule(this IServiceCollection services)
        {
            return services
                .AddScoped<IBenefitRepository, BenefitRepository>()
                .AddScoped<IBenefitService, BenefitService>();
        }
    }
}
