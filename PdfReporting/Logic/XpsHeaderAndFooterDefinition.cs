using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace PdfReporting.Logic
{
    public class XpsHeaderAndFooterDefinition
    {
        private readonly string _headerTemplateFilePath;
        private readonly string _footerTemplateFilePath;
        private readonly object _dataSourceObject;

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
        public Size PageSize { get; set; } = new Size(8.5 * 96.0, 11.0 * 96.0);

        /// <summary>
        /// Space reserved for the header in DIUs
        /// </summary>
        public double HeaderHeight
        {
            get => GetHeightOfTemplate(this._headerTemplateFilePath, this._dataSourceObject);
        }

        /// <summary>
        /// Space reserved for the footer in DIUs
        /// </summary>
        public double FooterHeight
        {
            get => GetHeightOfTemplate(this._footerTemplateFilePath, this._dataSourceObject);
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

        public XpsHeaderAndFooterDefinition(string headerTemplateFilePath, string footerTemplateFilePath, object dataSourceObject)
        {
            this._headerTemplateFilePath = headerTemplateFilePath;
            this._footerTemplateFilePath = footerTemplateFilePath;
            this._dataSourceObject = dataSourceObject;
        }

        private Visual GetVisualFromTemplate<T>(string templateFilePath, T dataSourceObject)
        {
            if(templateFilePath == null)
                return null;

            var managedFlowDocument = ManagedFlowDocument.GetFlowDocumentFrom(templateFilePath, dataSourceObject);
            managedFlowDocument = managedFlowDocument.GetCopy();
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
