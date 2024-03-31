using System.ComponentModel.DataAnnotations;

namespace BenefitsApp.Core.Models
{
    public class Benefit
    {
        [Key]
        public required string Code { get; set; }

        public required string Name { get; set; }

        public decimal RetailPrice { get; set; }

        public decimal DealerPrice { get; set; }

        public decimal SpecialPrice { get; set; }

        public int WarrantyPeriod { get; set; }

        public string? Note { get; set; }

        public required Category Category { get; set; }
    }
}
