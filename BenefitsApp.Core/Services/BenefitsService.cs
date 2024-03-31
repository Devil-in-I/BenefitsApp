using BenefitsApp.Core.Data;
using BenefitsApp.Core.Models;
using Microsoft.EntityFrameworkCore;
using System.Runtime.CompilerServices;

namespace BenefitsApp.Core.Services
{
    public class BenefitsService(BenefitsDbContext context) : IBenefitsService
    {
        public async Task<IList<Benefit>> GetAllBenefitsAsync(CancellationToken cancellationToken = default)
        {
            return await context.Benefits
                .Include(b => b.Category)
                .ToListAsync(cancellationToken);
        }

        public async IAsyncEnumerable<Benefit> StreamBenefitsWithCategory([EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            await foreach (var benefit in context.Benefits
                                                .Include(b => b.Category)
                                                .AsAsyncEnumerable()
                                                .WithCancellation(cancellationToken))
            {
                yield return benefit;
            }
        }
    }
}
