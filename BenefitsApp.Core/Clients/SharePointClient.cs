using PnP.Core.Auth.Services;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BenefitsApp.Core.Clients
{
    public class SharePointClient : ISharePointClient
    {
        private readonly PnPContextFactory _pnpContextFactory;
        private readonly IAuthenticationProviderFactory _authProviderFactory;

        public SharePointClient(PnPContextFactory pnpContextFactory, IAuthenticationProviderFactory authProviderFactory)
        {
            _pnpContextFactory = pnpContextFactory;
            _authProviderFactory = authProviderFactory;
        }

        public async Task<PnPContext> GetContextAsync(string siteUrl, string clientId, string clientSecret, string tenantId)
        {
            // Create an instance of the AuthenticationProvider
            var authProvider = _authProviderFactory.CreateWithClientCredentials(clientId, clientSecret, new Guid(tenantId));

            // Use the PnPContextFactory to get a PnPContext
            var context = await _pnpContextFactory.CreateAsync(new Uri(siteUrl), authProvider);

            return context;
        }
    }

}
