using BenefitsApp.Core.Models;
using Microsoft.Extensions.Options;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;

namespace BenefitsApp.Core.Services
{
    public class SharePointService : ISharePointService
    {
        private readonly IPnPContextFactory _pnpContextFactory;
        private readonly SharePointCredentialsOptions _sharePointCredentialsOptions;
        private readonly SharepointIDsOptions _sharePointIdsOptions;

        public SharePointService(
            IPnPContextFactory pnpContextFactory,
            IOptions<SharePointCredentialsOptions> sharePointCredentialsOptions,
            IOptions<SharepointIDsOptions> sharePointIdsOptions
            )
        {
            _pnpContextFactory = pnpContextFactory;
            _sharePointCredentialsOptions = sharePointCredentialsOptions.Value;
            _sharePointIdsOptions = sharePointIdsOptions.Value;
        }

        public async Task<PnPContext> GetContextAsync()
        {
            // Use the PnPContextFactory to get a PnPContext
            return await _pnpContextFactory.CreateAsync(new Uri(_sharePointCredentialsOptions.SiteUrl));
        }

        public async Task<IFieldCollection> GetAvailableFieldsAsync()
        {
            using var ctx = await GetContextAsync();

            await ctx.Web.LoadAsync(x => x.AvailableFields);

            return ctx.Web.AvailableFields;
        }

        public async Task<string> GetSiteTitleAsync()
        {
            using var context = await GetContextAsync();

            await context.Web.LoadAsync(x => x.Title);

            return context.Web.Title;
        }

        public async Task<string> GetSiteTitleAsync(PnPContext context)
        {
            await context.Web.LoadAsync(x => x.Title);

            return context.Web.Title;
        }

        public async Task<IFolderCollection> GetAllFoldersAsync()
        {
            using var context = await GetContextAsync();

            await context.Web.LoadAsync(x => x.Folders);

            return context.Web.Folders;
        }

        public async Task<IFolder> GetSharedDocumentsFolderByIdAsync()
        {
            using var context = await GetContextAsync();

            Guid sharedDocumentsFolderId = new Guid(_sharePointIdsOptions.SharedDocumentsFolderId);

            return await context.Web.GetFolderByIdAsync(sharedDocumentsFolderId);
        }

        public async Task<IFolder> GetBenefitsFolderByIdAsync()
        {
            using var context = await GetContextAsync();

            Guid benefitsFolderId = new Guid(_sharePointIdsOptions.BenefitsFolderId);

            return await context.Web.GetFolderByIdAsync(benefitsFolderId);
        }

        public async Task<IFolder> GetShopKzBenefitsFolderByIdAsync()
        {
            using var context = await GetContextAsync();

            Guid shopKzBenefitsFolderId = new(_sharePointIdsOptions.ShopKzBenefitsFolderId);

            return await context.Web.GetFolderByIdAsync(shopKzBenefitsFolderId);
        }

        public async Task<IFile> GetShopKzBenefitsXlsxDocByIdAsync()
        {
            using var context = await GetContextAsync();

            Guid ShopKzBenefitsXlsxDocId = new Guid(_sharePointIdsOptions.ShopKzBenefitsExcelDocumentId);

            return await context.Web.GetFileByIdAsync(ShopKzBenefitsXlsxDocId);
        }
    }

}
