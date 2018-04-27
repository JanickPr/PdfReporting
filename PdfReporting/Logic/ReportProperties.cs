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

        public Orientation ReportOrientation
        {
            get; set;
        }

        public ReportProperties(string templateFolderPath, PageNumberSettings pageNumberSettings, string outputDirectory, Orientation reportOrientation)
        {
            this.TemplateFolderPath = templateFolderPath;
            this.PageNumberSettings = pageNumberSettings;
            this.OutputDirectory = outputDirectory;
            this.ReportOrientation = reportOrientation;
        }
    }
}
