using PnP.Core.Model.SharePoint;
using PnP.Core.Services;

namespace BenefitsApp.Core.Services
{
    public interface ISharePointService
    {
        public Task<PnPContext> GetContextAsync();
        public Task<string> GetSiteTitleAsync();
        public Task<string> GetSiteTitleAsync(PnPContext context);
        public Task<IFieldCollection> GetAvailableFieldsAsync();
        public Task<IFolderCollection> GetAllFoldersAsync();
        public Task<IFolder> GetSharedDocumentsFolderByIdAsync();
        public Task<IFolder> GetBenefitsFolderByIdAsync();
        public Task<IFolder> GetShopKzBenefitsFolderByIdAsync();
        public Task<IFile> GetShopKzBenefitsXlsxDocByIdAsync();

    }
}
