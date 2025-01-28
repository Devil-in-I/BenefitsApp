//using BenefitsApp.Core.Data;
//using BenefitsApp.Core.Models;
//using Microsoft.EntityFrameworkCore;

//namespace BenefitsApp.Core.Services
//{
//    public class CategoriesService(BenefitsDbContext context) : ICategoryService
//    {
//        public async Task<IList<Category>> GetAllCategories(CancellationToken cancellationToken = default)
//        {
//            return await context.Categories
//                .ToListAsync(cancellationToken);
//        }

//        public async Task<IList<Category>> GetAllCategoriesWithBenefits(CancellationToken cancellationToken = default)
//        {
//            return await context.Categories
//                .Include(c => c.Benefits)
//                .ToListAsync(cancellationToken);
//        }
//    }
//}
