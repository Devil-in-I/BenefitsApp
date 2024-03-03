using PnP.Core.Model.SharePoint;
using PnP.Core.Services;

namespace BenefitsApp.Core.Clients
{
    public interface ISharePointService
    {
        public Task<PnPContext> GetContextAsync();
        public Task<string> GetBenefitsAsync();
    }
}
