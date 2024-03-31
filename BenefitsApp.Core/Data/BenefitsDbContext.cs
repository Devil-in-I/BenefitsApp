using BenefitsApp.Core.Models;
using Microsoft.EntityFrameworkCore;

namespace BenefitsApp.Core.Data
{
    /// <summary>
    /// Read-only database context to fetch Benefits from database.
    /// </summary>
    public class BenefitsDbContext : DbContext
    {
        public BenefitsDbContext(DbContextOptions options) : base(options)
        {
        }

        public DbSet<Benefit> Benefits { get; set; }
        public DbSet<Category> Categories { get; set; }

        public override int SaveChanges()
        {
            throw new InvalidOperationException("This context is read-only.");
        }

    }
}
