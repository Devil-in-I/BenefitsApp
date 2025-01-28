using BenefitsApp.Core.Data;
using BenefitsApp.Core.Models;
using BenefitsApp.Core.Repositories.Interfaces;

namespace BenefitsApp.Core.Repositories;

public class BenefitRepository(BenefitsDbContext context) : IBenefitRepository
{
    public IQueryable<Benefit> GetAllBenefits()
    {
        return context.Benefits.AsQueryable();
    }
}
