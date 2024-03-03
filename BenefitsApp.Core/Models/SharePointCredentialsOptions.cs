using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BenefitsApp.Core.Models
{
    public class SharePointCredentialsOptions
    {
        public const string SharePointCredentials = "SharePointCredentials";

        public string SiteUrl { get; set; } = string.Empty;
        public string ClientId { get; set; } = string.Empty;
        public string ClientSecret { get; set; } = string.Empty;
    }
}
