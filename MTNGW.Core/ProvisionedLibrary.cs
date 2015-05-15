using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTNGW.Core
{
    [Serializable]
    public class ProvisionedLibrary
    {
        public string Title { get; set; }
        public int TemplateType { get; set; }
        public List<ProvisionedContentType> ContentTypes { get; set; }

        public ProvisionedLibrary() { }
    }
}
