using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTNGW.Core
{
    [Serializable]
    public class SiteProvisioningSiteRequestType
    {
        public string TypeName { get; set; }
        public string TypeSharePointSiteId { get; set; }
        public string ThemeColourName { get; set; }
        public string DefaultSitePolicyName { get; set; }
        public int SiteStorageQuota { get; set; }
        public int SiteResourceLevel { get; set; }
        public string CustomCssUrl { get; set; }
        public string CustomLogoUrl { get; set; }
        public List<ProvisionedLibrary> Libraries { get; set; }

        public SiteProvisioningSiteRequestType() { }

    }
}
