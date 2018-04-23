using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfReporting.Logic
{
    public class ReportProperties
    {
        public string TemplateFolderPath
        {
            get; set;
        }

        public PageNumberSettings PageNumberSettings
        {
            get; set;
        }

        public string OutputDirectory
        {
            get; set;
        } 

        public ReportProperties(string templateFolderPath, PageNumberSettings pageNumberSettings, string outputDirectory)
        {
            TemplateFolderPath = templateFolderPath;
            PageNumberSettings = pageNumberSettings;
            OutputDirectory = outputDirectory;
        }
    }
}
