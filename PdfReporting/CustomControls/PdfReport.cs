using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PdfReporting.CustomControls
{
    /// <summary>
    /// Represents the blueprint for the Report.
    /// </summary>
    public class PdfReport : ContentControl
    {
        #region Fields

        private Orientation _orientation;
        private ContentControl _header;
        private ContentControl _footer;
        private bool _showPageNumbers;

        #endregion

        #region Properties

        /// <summary>
        /// Defines the orientation of all reportpages.
        /// </summary>
        public Orientation Orientation { get => _orientation; set => _orientation = value; }

        /// <summary>
        /// The headercontrol.
        /// </summary>
        public ContentControl Header { get => _header; set => _header = value; }

        /// <summary>
        /// The footercontrol.
        /// </summary>
        public ContentControl Footer { get => _footer; set => _footer = value; }

        /// <summary>
        /// Indicates whether the page numbers should be shown.
        /// </summary>
        public bool ShowPageNumbers { get => _showPageNumbers; set => _showPageNumbers = value; }


        #endregion

        #region Constructors

        static PdfReport()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(PdfReport), new FrameworkPropertyMetadata(typeof(PdfReport)));
        }

        #endregion
    }
}
