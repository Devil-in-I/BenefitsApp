using BenefitsApp.Core.Models;

namespace BenefitsApp.Core.Services
{
    public interface IExcelService
    {
        ValueTask<IEnumerable<Product>> GetProductsFromExcel(Stream excelStream);
        ValueTask<IEnumerable<Product>> GetProductsFromExcel(FileInfo file);
        IAsyncEnumerable<Product> GetProductsFromExcel2(string path);
    }
}
