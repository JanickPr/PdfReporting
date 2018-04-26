using System.Windows;
using System.Windows.Media;

namespace PdfReporting.Logic
{
    public class EditablePage : ContainerVisual
    {
        public void AddFooterAndHeader(XpsHeaderAndFooterDefinition xpsHeaderAndFooterDefinition)
        {
            this.AddVisualAtTopOfPage(xpsHeaderAndFooterDefinition.HeaderVisual);
            this.AddVisualAt(xpsHeaderAndFooterDefinition.FooterOffsetY, xpsHeaderAndFooterDefinition.FooterVisual);
        }

        private void AddVisualAtTopOfPage(Visual visual)
        {
            ContainerVisual containerVisual = GetContainerVisualWith(0);
            containerVisual.Children.Add(visual);
            this.Children.Add(containerVisual);
        }

        public void AddVisualAt(double offsetY, Visual visual)
        {
            if(visual == null)
                return;
            ContainerVisual containerVisual = GetContainerVisualWith(offsetY);
            containerVisual.Children.Add(visual);
            this.Children.Add(containerVisual);
        }

        private ContainerVisual GetContainerVisualWith(double offsetY)
        {
            var contentVisual = new ContainerVisual
            {
                Transform = new TranslateTransform(0, offsetY)
            };
            return contentVisual;
        }

        public void AddPageNumber(int pageNumber, int totalPages, PageNumberSettings pageNumberSettings)
        {
            var drawingVisual = new DrawingVisual();
            using(DrawingContext ctx = drawingVisual.RenderOpen())
            {
                FormattedText text = GetPageNumberText(pageNumberSettings, pageNumber, totalPages);
                ctx.DrawText(text, pageNumberSettings.Position);
            }
            this.Children.Add(drawingVisual);
        }

        private FormattedText GetPageNumberText(PageNumberSettings pageNumberSettings, int pageNumber, int totalPages)
        {
            if(pageNumberSettings.PageOfAllPagesNotation)
                return GetPageNumberTextEnhanced(pageNumberSettings, pageNumber, totalPages);
            else
                return GetPageNumberTextNormal(pageNumberSettings, pageNumber);
        }

        private FormattedText GetPageNumberTextEnhanced(PageNumberSettings pageNumberSettings, int pageNumber, int totalPages)
        {
            return new FormattedText(pageNumberSettings.PagePrefix + " " + (pageNumber) + " von " + totalPages,
                   System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                   pageNumberSettings.FontFamily, pageNumberSettings.FontSize, pageNumberSettings.FontBrush);
        }

        private FormattedText GetPageNumberTextNormal(PageNumberSettings pageNumberSettings, int pageNumber)
        {
            return new FormattedText(pageNumberSettings.PagePrefix + " " + (pageNumber + 1),
                System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                pageNumberSettings.FontFamily, pageNumberSettings.FontSize, pageNumberSettings.FontBrush);
        }
    }
}
