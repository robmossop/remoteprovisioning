using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTNGW.Core
{
    [Serializable]
    public class ProvisionedPage
    {
        public string Url { get; set; }
        public string Title { get; set; }
        public List<ProvisionedWebPart> PageWebParts { get; set; }

        public ProvisionedPage() { }
    }
}
