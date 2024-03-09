namespace BenefitsApp.Core.Models
{
    public class Product // TODO: To think on whether or not it's better to use Record.
    {
        public required string Code { get; set; }
        public required string Name { get; set; }
        public decimal RetailPrice { get; set; }
        public decimal DealerPrice { get; set; }
        public decimal SpecialPrice { get; set; }
        public int WarrantyPeriod { get; set; }
        public string? Note { get; set; }
        public required string Category { get; set; }
    }
}
