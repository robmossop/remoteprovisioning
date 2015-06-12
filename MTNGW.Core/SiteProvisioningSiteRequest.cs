using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTNGW.Core
{
    public class SiteProvisioningSiteRequest
    {
        public string Title { get; set; }
        public int SpItemId { get; set; }
        public string Description { get; set; }
        public string Url { get; set; }
        public string HostUrl { get; set; }
        public FieldUserValue[] SiteAdmins { get; set; }
        public FieldUserValue[] SiteOwners { get; set; }
        public FieldUserValue[] SiteMembers { get; set; }
        public FieldUserValue[] SiteVisitors { get; set; }
        public string SiteTemplate { get; set; }
        public ClientContext ClientCtx { get; set; }
        public string SiteId { get; set; }
        public string SiteStatus { get; set; }
        public int SiteQuota { get; set; }

        public SiteProvisioningSiteRequest() { }

        public string CreateSite() { return string.Empty; }

        public string DeleteSite() { return string.Empty; }
    }
}
