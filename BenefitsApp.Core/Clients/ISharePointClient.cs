using PnP.Core.Services;

namespace BenefitsApp.Core.Clients
{
    public interface ISharePointClient
    {
        public Task<PnPContext> GetContextAsync(string siteUrl, string clientId, string clientSecret, string tenantId);
    }
}
