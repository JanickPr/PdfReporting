using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Runtime.Serialization;
using PdfReporting.Logic;

namespace PdfReporting.Logic
{
    public class XpsHeaderAndFooterDefinition
    {

        #region Page sizes

        private Thickness _margins;// = new Thickness(96); // Default: 1" margins
        private Size _pageSize = new Size(793.5987, 1122.3987); // Default: A4
        private bool _repeatTableHeaders = true;
        private double _headerHeight;
        private double _footerHeight;
        private readonly string _headerTemplateFilePath;
        private readonly string _footerTemplateFilePath;
        private readonly string _bodyTemplateFilePath;
        private readonly object _dataSourceObject;

        public Visual HeaderVisual => GetVisualFromTemplate(_headerTemplateFilePath, _dataSourceObject);
        
        public Visual FooterVisual => GetVisualFromTemplate(_footerTemplateFilePath, _dataSourceObject);

        /// <summary>
        /// PageSize in DIUs
        /// </summary>
        public Size PageSize
        {
            get
            {
                return _pageSize;
            }
            set
            {
                _pageSize = value;
            }
        }

        /// <summary>
        /// Margins
        /// </summary>
        public Thickness Margins
        {
            get
            {
                return _margins;
            }
            set
            {
                _margins = value;
            }
        }



        /// <summary>
        /// Space reserved for the header in DIUs
        /// </summary>
        public double HeaderHeight
        {
            get
            {
                return GetHeightOfTemplate(_headerTemplateFilePath, _dataSourceObject);
            }
        }

        /// <summary>
        /// Space reserved for the footer in DIUs
        /// </summary>
        public double FooterHeight
        {
            get
            {
                return GetHeightOfTemplate(_footerTemplateFilePath, _dataSourceObject);
            }
        }

        public double ContentHeight
        {
            get
            {
                return PageSize.Height - FooterHeight - HeaderHeight;
            }
        }

        #endregion

        ///<summary>
        /// Should table headers automatically repeat?
        ///</summary>
        public bool RepeatTableHeaders
        {
            get
            {
                return _repeatTableHeaders;
            }
            set
            {
                _repeatTableHeaders = value;
            }
        }

        #region Some convenient helper properties

        internal Size ContentSize
        {
            get
            {
                return new Size(
                   PageSize.Width - (Margins.Left + Margins.Right),
                   PageSize.Height - (Margins.Top + Margins.Bottom + HeaderHeight + FooterHeight)
                );
            }
        }

        internal Point ContentOrigin
        {
            get
            {
                return new Point(
                    Margins.Left,
                    Margins.Top + HeaderRect.Height
                );
            }
        }

        internal Rect HeaderRect
        {
            get
            {
                return new Rect(
                    Margins.Left, Margins.Top,
                    ContentSize.Width, HeaderHeight
                );
            }
        }

        internal Rect FooterRect
        {
            get
            {
                return new Rect(
                    Margins.Left, ContentOrigin.Y + ContentSize.Height,
                    ContentSize.Width, FooterHeight
                );
            }
        }

        public XpsHeaderAndFooterDefinition(String headerTemplateFilePath, String footerTemplateFilePath, object dataSourceObject)
        {
            this._headerTemplateFilePath = headerTemplateFilePath;
            this._footerTemplateFilePath = footerTemplateFilePath;
            this._dataSourceObject = dataSourceObject;
        }

        private Visual GetVisualFromTemplate<T>(string templateFilePath, T dataSourceObject)
        {
            if (templateFilePath == null)
                return null;
            
            ManagedFlowDocument managedFlowDocument = ManagedFlowDocument.GetFlowDocumentFrom(templateFilePath, dataSourceObject);
            managedFlowDocument = managedFlowDocument.GetCopy();
            SetDimensonsFor(ref managedFlowDocument);
            var visual = managedFlowDocument.GetVisualOfPage(0);
            return visual;
        }

        private void SetDimensonsFor(ref ManagedFlowDocument managedFlowDocument)
        {
            managedFlowDocument.ColumnWidth = double.MaxValue; // Prevent columns
            managedFlowDocument.PageWidth = this.ContentSize.Width;
        }

        private Double GetHeightOfTemplate<T>(string templateFilePath, T dataSourceObject)
        {
            if(templateFilePath == null)
                return 0;

            ManagedFlowDocument managedFlowDocument = ManagedFlowDocument.GetFlowDocumentFrom(templateFilePath, dataSourceObject);
            DocumentPage page =  managedFlowDocument.GetPage(0);
            return page.Size.Height;
        }

        #endregion

    }
}
