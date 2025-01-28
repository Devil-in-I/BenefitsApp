using BenefitsApp.Core.Models;

namespace BenefitsApp.Core.Services
{
    public interface IBenefitService
    {
        public Task<IReadOnlyList<Benefit>> GetAllBenefitsAsync(CancellationToken cancellationToken = default);
        long GetBenefitsCount();
    }
}
