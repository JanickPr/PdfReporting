using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace PdfReporting.Logic
{
    public class ReportContentDefinition
    {
        private readonly string _headerTemplateFilePath;
        private readonly string _footerTemplateFilePath;
        private readonly object _dataSourceObject;
        private readonly ReportProperties _reportProperties;

        private Visual _headerVisual, _footerVisual;
        private double _headerHeight, _footerHeight;
        private Size _pageSize;

        public Visual HeaderVisual
        {
            get => this.GetVisualFromTemplate(this._headerTemplateFilePath, this._dataSourceObject);
        }

        public Visual FooterVisual
        {
            get => this.GetVisualFromTemplate(this._footerTemplateFilePath, this._dataSourceObject);
        }

        /// <summary>
        /// PageSize in DIUs
        /// </summary>
        public Size PageSize
        {
            get => this._pageSize == default ? (this._pageSize = this._reportProperties.ReportOrientation == Orientation.Vertical ? new Size(793.5987, 1122.3987) : new Size(1122.3987, 793.5987)) : _pageSize;
        }

        /// <summary>
        /// Space reserved for the header in DIUs
        /// </summary>
        public double HeaderHeight
        {
            get => this._headerHeight == 0 ? this._headerHeight = GetHeightOfTemplate(this._headerTemplateFilePath, this._dataSourceObject) : this._headerHeight;
        }

        /// <summary>
        /// Space reserved for the footer in DIUs
        /// </summary>
        public double FooterHeight
        {
            get => this._footerHeight == 0 ? this._footerHeight = GetHeightOfTemplate(this._footerTemplateFilePath, this._dataSourceObject) : this._footerHeight;
        }

        public double FooterOffsetY
        {
            get => this.PageSize.Height - this.FooterHeight;
        }

        public double ContentHeight
        {
            get => this.PageSize.Height - this.FooterHeight - this.HeaderHeight;
        }

        internal Size ContentSize
        {
            get
            {
                return new Size(this.PageSize.Width,
                   this.PageSize.Height - (this.HeaderHeight + this.FooterHeight)
                );
            }
        }

        public ReportContentDefinition(string headerTemplateFilePath, string footerTemplateFilePath, object dataSourceObject, ReportProperties reportProperties)
        {
            this._headerTemplateFilePath = headerTemplateFilePath;
            this._footerTemplateFilePath = footerTemplateFilePath;
            this._dataSourceObject = dataSourceObject;
            this._reportProperties = reportProperties;
        }

        private Visual GetVisualFromTemplate<T>(string templateFilePath, T dataSourceObject)
        {
            if(templateFilePath == null)
                return null;

            var managedFlowDocument = ManagedFlowDocument.GetFlowDocumentFrom(templateFilePath, dataSourceObject);  
            SetDimensionsOf(managedFlowDocument);
            return managedFlowDocument.GetVisualOfPage(0);
        }

        private double GetHeightOfTemplate<T>(string templateFilePath, T dataSourceObject)
        {
            if(templateFilePath == null)
                return 0;

            var managedFlowDocument = ManagedFlowDocument.GetFlowDocumentFrom(templateFilePath, dataSourceObject);
            DocumentPage page = managedFlowDocument.GetPage(0);
            return page.Size.Height;
        }

        public ManagedFlowDocument SetDimensionsOf(ManagedFlowDocument managedFlowDocument)
        {
            managedFlowDocument.ColumnWidth = double.MaxValue; // Prevent columns
            managedFlowDocument.PageWidth = this.ContentSize.Width;
            managedFlowDocument.PageHeight = this.ContentHeight;
            managedFlowDocument.PagePadding = new Thickness(0);
            return managedFlowDocument;
        }
    }
}
