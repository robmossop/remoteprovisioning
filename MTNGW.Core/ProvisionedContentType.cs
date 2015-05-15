using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTNGW.Core
{
    [Serializable]
    public class ProvisionedContentType
    {
        public string ContentTypeId { get; set; }
        public string Name { get; set; }

        public ProvisionedContentType() { }
    }
}
