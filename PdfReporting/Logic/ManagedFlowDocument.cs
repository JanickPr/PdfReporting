using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace PdfReporting.Logic
{
    public class ManagedFlowDocument : FlowDocument
    {
        public new void RemoveLogicalChild(object child)
        {
            base.RemoveLogicalChild(child);
        }
    }
}
