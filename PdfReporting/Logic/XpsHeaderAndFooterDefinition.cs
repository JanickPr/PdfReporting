using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace PdfReporting.Logic
{
    public class XpsHeaderAndFooterDefinition {

        #region Page sizes

        private Thickness _Margins = new Thickness(96); // Default: 1" margins
        private Size _PageSize = new Size(793.5987, 1122.3987); // Default: A4
        private bool _RepeatTableHeaders = true;
        private double _HeaderHeight;
        private double _FooterHeight;

        public Visual HeaderVisual
        {
            get; set;
        }

        public Visual FooterVisual
        {
            get; set;
        }

        /// <summary>
        /// PageSize in DIUs
        /// </summary>
        public Size PageSize {
			get { return _PageSize; }
			set { _PageSize = value; }
		}

		/// <summary>
		/// Margins
		/// </summary>
		public Thickness Margins {
			get { return _Margins; }
			set { _Margins = value; }
		}
			


		/// <summary>
		/// Space reserved for the header in DIUs
		/// </summary>
		public double HeaderHeight {
			get { return _HeaderHeight; }
			set { _HeaderHeight = value; }
		}

		/// <summary>
		/// Space reserved for the footer in DIUs
		/// </summary>
		public double FooterHeight {
			get { return _FooterHeight; }
			set { _FooterHeight = value; }
		}

		#endregion

        ///<summary>
        /// Should table headers automatically repeat?
        ///</summary>
        public bool RepeatTableHeaders {
			get { return _RepeatTableHeaders; }
			set { _RepeatTableHeaders = value; }
		}

        #region Some convenient helper properties

        internal Size ContentSize
        {
            get
            {
                return new Size(
                    Margins.Left + Margins.Right,
                    Margins.Top + Margins.Bottom + HeaderHeight + FooterHeight
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
            CreateHeaderAndFooterVisualFrom(headerTemplateFilePath, footerTemplateFilePath, dataSourceObject);
        }

        private void CreateHeaderAndFooterVisualFrom(string headerTemplateFilePath, string footerTemplateFilePath, object dataSourceObject)
        {
            HeaderVisual = GetVisualFromTemplate(headerTemplateFilePath, dataSourceObject);
            FooterVisual = GetVisualFromTemplate(footerTemplateFilePath, dataSourceObject);
        }

        private Visual GetVisualFromTemplate(string templateFilePath, object dataSourceObject)
        {
            if (templateFilePath == null)
                return null;

            FlowDocument flowDocument = GetFlowDocumentFromTemplate(templateFilePath);
            flowDocument.SetUpDataContext(dataSourceObject);
            return ((IDocumentPaginatorSource)flowDocument).DocumentPaginator.GetPage(0).Visual;
        }

        private FlowDocument GetFlowDocumentFromTemplate(string templateFilePath)
        {
            FlowDocument flowDocument = new FlowDocument();
            flowDocument = flowDocument.LoadFromTemplate(templateFilePath);
            return flowDocument;
        }

        #endregion

    }
}
