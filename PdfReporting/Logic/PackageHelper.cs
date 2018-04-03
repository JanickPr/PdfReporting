using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfReporting.Logic
{
    public static class PackageHelper
    {
        public static Package OpenPackageInMemoryStream()
        {
            MemoryStream memoryStream = new MemoryStream();
            return Package.Open(memoryStream, FileMode.Create);
        }
    }
}
