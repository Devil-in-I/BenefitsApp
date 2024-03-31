using BenefitsApp.Core.Models;

namespace BenefitsApp.Core.Services
{
    public interface IBenefitsService
    {
        public Task<IList<Benefit>> GetAllBenefitsAsync(CancellationToken cancellationToken = default);
        public IAsyncEnumerable<Benefit> StreamBenefitsWithCategory(CancellationToken cancellationToken = default);
    }
}
