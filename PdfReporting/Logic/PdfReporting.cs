using PdfReporting.CustomControls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace PdfReporting.Logic
{
    public static class PdfReporting
    {
        #region Fields

        private static double standartPageHeight = 1123.2;
        private static double standartPageWidth = 796.8;
        private static int length;

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
        /// <param name="frameworkElement"></param>
        /// <returns></returns>
        private static List<FixedPage> DivideReportIntoPages(FrameworkElement frameworkElement)
        {
            if (frameworkElement == null)
            {
                throw new ArgumentNullException(nameof(frameworkElement));
            }

            FrameworkElement pageContent;
            FrameworkElement currentFrameworkElement = frameworkElement;

            int visualTreeLevel = 0;
            double cumulatedHeight = 0;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(currentFrameworkElement); i++)
            {
                double elementHeightContentIncluded = GetFrameworkElementHeightIncludingContent(currentFrameworkElement);

                if (cumulatedHeight + elementHeightContentIncluded < standartPageHeight)
                {
                    //The full Element is added to pageContent at the correct position.
                    AddToElement(currentFrameworkElement, ref pageContent);

                    if (IsSameOrSubclass(typeof(Panel), currentFrameworkElement.GetType()) &&
                        PanelHasVerticalOrientedItems((Panel)currentFrameworkElement))
                    {
                        //We must iterate through all items because the Panle can be distributet over several pages.
                    }
                }

                var pagesList = new List<FixedPage>();

                return pagesList;
            }
        }

        
        private static void AddToElement(FrameworkElement elementToBeAdded, ref FrameworkElement parentElement)
        {
            if (parentElement.GetType() == typeof(ContentControl))
                ((ContentControl)parentElement).Content = elementToBeAdded;
            else if (parentElement.GetType() == typeof(Panel))
                ((Panel)parentElement).Children.Add(elementToBeAdded);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="panel"></param>
        /// <returns></returns>
        private static bool PanelHasVerticalOrientedItems(Panel panel)
        {
            if (panel.GetType() == typeof(Grid))
            {
                Grid grid = (Grid)panel;
                return grid.RowDefinitions.Count > 0;
            }

            return panel.LogicalOrientationPublic == Orientation.Vertical;

            //Here we must check for other paneltypes e.g. Canvas or WrapPanel
        }

        /// <summary>
        /// Returns the height that is used by the given frameworkelement with the content included.
        /// </summary>
        /// <param name="frameworkElement"></param>
        /// <returns></returns>
        private static double GetFrameworkElementHeightIncludingContent(FrameworkElement frameworkElement)
        {
            double height = GetFrameworkElementHeightExcludingContent(frameworkElement);

            height += frameworkElement.ActualHeight;

            return height;
        }

        /// <summary>
        /// Returns the height that is used by the given frameworkelement without the content.
        /// </summary>
        /// <param name="frameworkElement"></param>
        /// <returns></returns>
        private static double GetFrameworkElementHeightExcludingContent(FrameworkElement frameworkElement)
        {
            double height = 0;

            height += GetHeightFromSpacingproperty(frameworkElement.Margin);

            if (frameworkElement.GetType().GetProperty("Padding") != null)
            {
                var padding = frameworkElement.GetType().GetProperty("Padding").GetValue(frameworkElement);
                height += GetHeightFromSpacingproperty((Thickness)padding);
            }

            if (frameworkElement.GetType().GetProperty("BorderThickness") != null)
            {
                var borderThickness = frameworkElement.GetType().GetProperty("BorderThickness").GetValue(frameworkElement);
                height += GetHeightFromSpacingproperty((Thickness)borderThickness);
            }

            return height;
        }

        /// <summary>
        /// returns the vertical space that is used for spacing such as margin or padding.
        /// </summary>
        /// <param name="spacing"></param>
        /// <returns></returns>
        private static double GetHeightFromSpacingproperty(Thickness spacing)
        {
            return spacing.Top + spacing.Bottom;
        }

        /// <summary>
        /// Returns true if the given descendant type derives from the given base type
        /// or if the decendant type is equal to the base type.
        /// </summary>
        /// <param name="Base"></param>
        /// <param name="Descendant"></param>
        /// <returns></returns>
        private static bool IsSameOrSubclass(Type Base, Type Descendant)
        {
            return Descendant.IsSubclassOf(Base)
                   || Descendant == Base;
        }

        #endregion
    }
}
