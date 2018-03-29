using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Xps.Packaging;

namespace PdfReporting.Logic
{
    public class ManagedXpsDocument : XpsDocument
    {
        public Uri PackageUri { get; set; }

        public ManagedXpsDocument(Package package, CompressionOption compressionOption, string path) : base(package, compressionOption, path)
        {

        }
    }
}
