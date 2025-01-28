using BenefitsApp.Core.Models;
using Microsoft.EntityFrameworkCore;

namespace BenefitsApp.Core.Data
{
    /// <summary>
    /// Read-only database context to fetch Benefits from database.
    /// </summary>
    public class BenefitsDbContext(DbContextOptions options) : DbContext(options)
    {
        public DbSet<Benefit> Benefits { get; set; }
        //public DbSet<Category> Categories { get; set; }

        public override int SaveChanges() => throw new InvalidOperationException("This context is read-only.");

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<Benefit>();
        }
    }
}
