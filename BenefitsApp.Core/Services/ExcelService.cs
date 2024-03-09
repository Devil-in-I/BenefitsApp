using BenefitsApp.Core.Models;
using OfficeOpenXml;

namespace BenefitsApp.Core.Services
{
    public class ExcelService : IExcelService
    {
        public async ValueTask<IEnumerable<Product>> GetProductsFromExcel(Stream excelStream)
        {
            ArgumentNullException.ThrowIfNull(excelStream);

            var products = new List<Product>();
            var currentCategory = string.Empty;

            using (var excelPackage = new ExcelPackage(excelStream))
            {
                await excelPackage.LoadAsync(excelStream);

                var worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

                ArgumentNullException.ThrowIfNull(worksheet);

                int rowCount = worksheet.Dimension.Rows;

                // Starting from the 8th row to skip headers and category row
                for (int row = 8; row <= rowCount; row++)
                {
                    // If all 7 rows are merged that it is row containing category name.
                    if (worksheet.Cells[row, 1, row, 7].Merge)
                    {
                        // Getting category row.
                        currentCategory = worksheet.Cells[row, 1].Value?.ToString();
                    }
                    else
                    {
                        var product = new Product
                        {
                            Code = worksheet.Cells[row, 1].Value?.ToString(),
                            Name = worksheet.Cells[row, 2].Value?.ToString(),
                            RetailPrice = decimal.Parse(worksheet.Cells[row, 3].Value?.ToString() ?? "0"),
                            DealerPrice = decimal.Parse(worksheet.Cells[row, 4].Value?.ToString() ?? "0"),
                            SpecialPrice = decimal.Parse(worksheet.Cells[row, 5].Value?.ToString() ?? "0"),
                            WarrantyPeriod = int.Parse(worksheet.Cells[row, 6].Value?.ToString() ?? "0"),
                            Note = worksheet.Cells[row, 7].Value?.ToString(),
                            Category = currentCategory
                        };

                        products.Add(product);
                    }
                }
            }

            return products;
        }

        public async ValueTask<IEnumerable<Product>> GetProductsFromExcel(FileInfo file)
        {
            ArgumentNullException.ThrowIfNull(file);

            var products = new List<Product>();
            var currentCategory = string.Empty;

            using (var excelPackage = new ExcelPackage(file))
            {
                await excelPackage.LoadAsync(file);

                var worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

                ArgumentNullException.ThrowIfNull(worksheet);

                int rowCount = worksheet.Dimension.Rows;

                // Starting from the 8th row to skip headers and category row
                for (int row = 8; row <= rowCount; row++)
                {
                    // If all 7 rows are merged that it is row containing category name.
                    if (worksheet.Cells[row, 1, row, 7].Merge)
                    {
                        // Getting category row.
                        currentCategory = worksheet.Cells[row, 1].Value?.ToString();
                    }
                    else
                    {
                        var product = new Product
                        {
                            Code = worksheet.Cells[row, 1].Value?.ToString(),
                            Name = worksheet.Cells[row, 2].Value?.ToString(),
                            RetailPrice = decimal.Parse(worksheet.Cells[row, 3].Value?.ToString() ?? "0"),
                            DealerPrice = decimal.Parse(worksheet.Cells[row, 4].Value?.ToString() ?? "0"),
                            SpecialPrice = decimal.Parse(worksheet.Cells[row, 5].Value?.ToString() ?? "0"),
                            WarrantyPeriod = int.Parse(worksheet.Cells[row, 6].Value?.ToString() ?? "0"),
                            Note = worksheet.Cells[row, 7].Value?.ToString(),
                            Category = currentCategory
                        };

                        products.Add(product);
                    }
                }
            }

            return products;
        }
    }
}
