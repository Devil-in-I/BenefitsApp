using BenefitsApp.Core.Models;
using Microsoft.Extensions.Options;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;

namespace BenefitsApp.Core.Clients
{
    public class SharePointService : ISharePointService
    {
        private readonly IPnPContextFactory _pnpContextFactory;
        private readonly SharePointCredentialsOptions _options;

        public SharePointService(IPnPContextFactory pnpContextFactory, IOptions<SharePointCredentialsOptions> options)
        {
            _pnpContextFactory = pnpContextFactory;
            _options = options.Value;
        }

        public async Task<PnPContext> GetContextAsync()
        {
            // Use the PnPContextFactory to get a PnPContext
            return await _pnpContextFactory.CreateAsync(new Uri(_options.SiteUrl));
        }

        public async Task<string> GetBenefitsAsync()
        {
            using var context = await GetContextAsync();

            await context.Web.LoadAsync(x => x.Title);

            return context.Web.Title;
        }
    }

}
