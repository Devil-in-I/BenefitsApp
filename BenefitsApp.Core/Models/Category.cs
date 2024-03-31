using System.ComponentModel.DataAnnotations;

namespace BenefitsApp.Core.Models
{
    public class Category
    {
        [Key]
        public int Id { get; set; }

        public required string Name { get; set; }

        public ICollection<Benefit>? Benefits { get; set; }
    }
}
