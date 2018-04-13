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
    public class DocumentPaginatorWrapper : DocumentPaginator
    {
        private Size _pageSize, _margin;
        private readonly XpsHeaderAndFooterDefinition _xpsHeaderAndFooterDefinition;
        private DocumentPaginator _paginator;
        private Typeface _typeface;

        public override bool IsPageCountValid => _paginator.IsPageCountValid;

        public override int PageCount => _paginator.PageCount;

        public override Size PageSize
        {
            get
            {
                return _paginator.PageSize;
            }
            set
            {
                _paginator.PageSize = value;
            }
        }

        public override IDocumentPaginatorSource Source => _paginator.Source;

        public DocumentPaginatorWrapper(DocumentPaginator paginator, Size pageSize, Size margin, XpsHeaderAndFooterDefinition xpsHeaderAndFooterDefinition)
        {
            _pageSize = pageSize;
            _margin = margin;
            this._xpsHeaderAndFooterDefinition = xpsHeaderAndFooterDefinition;
            _paginator = paginator;

            _paginator.PageSize = new Size(_pageSize.Width - margin.Width * 2,
                                            _pageSize.Height - margin.Height * 2);
        }

        Rect Move(Rect rect)
        {
            if (rect.IsEmpty)
            {
                return rect;
            }
            else
            {
                return new Rect(rect.Left + _margin.Width, rect.Top + _margin.Height,
                                rect.Width, rect.Height);
            }
        }

        public override DocumentPage GetPage(int pageNumber)
        {
            FlowDocument doc = new FlowDocument();
            DocumentPage page = _paginator.GetPage(pageNumber);

            // Create a wrapper visual for transformation and add extras
            ContainerVisual newpage = new ContainerVisual();

            DrawingVisual title = new DrawingVisual();

            using (DrawingContext ctx = title.RenderOpen())
            {
                if (_typeface == null)
                {
                    _typeface = new Typeface("Times New Roman");
                }

                FormattedText text = new FormattedText("Page " + (pageNumber + 1),
                    System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    _typeface, 14, Brushes.Black);

                ctx.DrawText(text, new Point(0, -96 / 4)); // 1/4 inch above page content
            }

            DrawingVisual background = new DrawingVisual();

            using (DrawingContext ctx = background.RenderOpen())
            {
                ctx.DrawRectangle(new SolidColorBrush(Color.FromRgb(240, 240, 240)), null, page.ContentBox);
            }

            //newpage.Children.Add(background); // Scale down page and center
            newpage.Children.Add(_xpsHeaderAndFooterDefinition.HeaderVisual);

            ContainerVisual smallerPage = new ContainerVisual();
            smallerPage.Children.Add(page.Visual);
            smallerPage.Transform = new MatrixTransform(0.95, 0, 0, 0.95,
                0.025 * page.ContentBox.Width, 0.025 * page.ContentBox.Height);
            
            newpage.Children.Add(smallerPage);
            newpage.Children.Add(title);

            newpage.Transform = new TranslateTransform(_margin.Width, _margin.Height);

            return new DocumentPage(newpage, _pageSize, Move(page.BleedBox), Move(page.ContentBox));
        } 
        
        

    }
}
