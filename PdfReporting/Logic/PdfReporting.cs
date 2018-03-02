using PdfReporting.CustomControls;
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
    public static class PdfReporting
    {
        #region Fields

        private static double pageHight = 1123.2;
        private static double pageWidth = 796.8;

        #endregion

        #region Methods

        /// <summary>
        /// Sets the  datacontext of the PdfReportBlueprint to the dataSource-object
        /// and creates an .pdf-file in the outputDirectory
        /// </summary>
        /// <param name="pdfReport">The blueprint for the pdfreport.</param>
        /// <param name="dataSource">The datasource which gets binded to the blueprint.</param>
        /// <param name="outputDirectory">The outputdirectory where the .pdf-file is saved to.</param>
        /// <param name="pdfFileName">The name of the .pdf-file that is created.</param>
        public static void CreatePdfFromReport(PdfReport pdfReport, object dataSource, string outputDirectory, string pdfFileName)
        {
            pdfReport = SetDataContext(pdfReport, dataSource);
            
        }

        /// <summary>
        /// Sets the datacontext for the pdfReportBlueprint.
        /// </summary>
        /// <param name="pdfReport"></param>
        /// <param name="dataSource"></param>
        /// <returns></returns>
        private static PdfReport SetDataContext(PdfReport pdfReport, object dataSource)
        {
            pdfReport.DataContext = dataSource;
            return pdfReport;
        }

        /// <summary>
        /// Goes through the reports visualtree and divides the report into several pages.
        /// </summary>
        /// <param name="pdfReport"></param>
        /// <returns></returns>
        private static List<FixedPage> DivideReportIntoPages(PdfReport pdfReport)
        {
            var pagesList = new List<FixedPage>();
            double actualPageHeight = 0;
            int maxPageCount =  Convert.ToInt32(Math.Ceiling(pdfReport.ActualHeight / pageHight));

            for(int pageCounter = 0; pageCounter < maxPageCount)
            {

            }

            return pagesList;
        }

        #endregion
    }
}
