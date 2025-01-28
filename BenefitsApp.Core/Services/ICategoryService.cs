using BenefitsApp.Core.Models;

namespace BenefitsApp.Core.Services
{
    public interface ICategoryService
    {
        public Task<IList<Category>> GetAllCategories(CancellationToken cancellationToken = default);
        public Task<IList<Category>> GetAllCategoriesWithBenefits(CancellationToken cancellationToken = default);
    }
}
