using BenefitsApp.Core.Models;

namespace BenefitsApp.Core.Repositories.Interfaces;

public interface IBenefitRepository
{
    public IQueryable<Benefit> GetAllBenefits();
}
