using BenefitsApp.Core.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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

        public async IAsyncEnumerable<Product> GetProductsFromExcel2(string path)
        {
            ArgumentNullException.ThrowIfNull(path);
            
            var products = new List<Product>();
            var currentCategory = string.Empty;

            using (var spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart; // Получаем часть книги
                ArgumentNullException.ThrowIfNull(workbookPart);

                var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault() ?? throw new ArgumentException("Worksheet not found in the document.");
                var worksheet = worksheetPart.Worksheet;
                var sheetData = worksheet.GetFirstChild<SheetData>() ?? throw new ArgumentException("Sheet data not found.");

                int rowCount = sheetData.Elements<Row>().Count();
                var productNumber = 0;
                // Начинаем с 8-й строки, чтобы пропустить заголовки и строку с категорией
                for (int row = 7; row <= rowCount; row++)
                {
                    var currentRow = sheetData.Elements<Row>().ElementAt(row - 1);
                    var cellValues = currentRow.Elements<Cell>().Select(cell => GetCellValue(cell, workbookPart)).ToList();
                    Product product = default;
                    // Если все 7 столбцов объединены, то это строка с именем категории.
                    if (string.IsNullOrWhiteSpace(cellValues.ElementAtOrDefault(1)))
                    {
                        // Получаем название категории.
                        currentCategory = cellValues.ElementAtOrDefault(2);
                        continue;
                    }
                    else
                    {
                        await Task.Delay(1);
                        product = new Product
                        {
                            Code = cellValues.ElementAtOrDefault(1) ?? throw new ArgumentException("Code cell was null."),
                            Name = cellValues.ElementAtOrDefault(2) ?? throw new ArgumentException("Name cell was null."),
                            RetailPrice = decimal.Parse(cellValues.ElementAtOrDefault(3) ?? "0"),
                            DealerPrice = decimal.Parse(cellValues.ElementAtOrDefault(4) ?? "0"),
                            SpecialPrice = decimal.Parse(cellValues.ElementAtOrDefault(5) ?? "0"),
                            WarrantyPeriod = int.Parse(cellValues.ElementAtOrDefault(6) ?? "0"),
                            Note = cellValues.ElementAtOrDefault(7),
                            Category = currentCategory ?? throw new ArgumentException("Category was null.")
                        };

                        //products.Add(product);
                        productNumber++;
                        Console.WriteLine($"*********** Added product #{productNumber}: {product.Name} - CODE:{product.Code}");
                    }
                    yield return product;
                }
            }

            //return products;
        }


        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null || cell.CellValue == null)
                return null;

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringTablePart != null)
                {
                    value = sharedStringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }

            return value;
        }
    }
}
