using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BenefitsApp.Core.Models
{
    public class SharepointIDsOptions
    {
        public const string SharepointIds = "SharepointIds";

        public string SharedDocumentsFolderId { get; set; } = string.Empty;
        public string BenefitsFolderId { get; set; } = string.Empty;
        public string ShopKzBenefitsFolderId { get; set; } = string.Empty;
        public string ShopKzBenefitsExcelDocumentId { get; set; } = string.Empty;
    }
}
