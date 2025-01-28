using BenefitsApp.Core.Models;
using BenefitsApp.Core.Repositories.Interfaces;
using Microsoft.EntityFrameworkCore;

namespace BenefitsApp.Core.Services
{
    public class BenefitService(IBenefitRepository benefitRepository) : IBenefitService
    {
        public async Task<IReadOnlyList<Benefit>> GetAllBenefitsAsync(CancellationToken cancellationToken = default)
        {
            return await benefitRepository
                .GetAllBenefits()
                .Take(5)
                .ToListAsync(cancellationToken);
        }

        public long GetBenefitsCount()
        {
            return benefitRepository
                .GetAllBenefits()
                .Count();
        }
    }
}
