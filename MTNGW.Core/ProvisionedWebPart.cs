using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTNGW.Core
{
    [Serializable]
    public class ProvisionedWebPart
    {
        public string Title { get; set; }
        public string TitleUrl { get; set; }
        public string WebPartXml { get; set; }

        public ProvisionedWebPart() { }
    }
}
