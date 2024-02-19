namespace BenefitsApp.Core.Clients
{
    public interface ISharePointClient
    {
        public Stream GetFile(string relativeUrl);
    }
}
