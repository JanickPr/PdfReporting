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
        private Size _PageSize, _Margin;
        private DocumentPaginator _Paginator;
        private Typeface _Typeface;

        public override bool IsPageCountValid => _Paginator.IsPageCountValid;

        public override int PageCount => _Paginator.PageCount;

        public override Size PageSize
        {
            get
            {
                return _Paginator.PageSize;
            }
            set
            {
                _Paginator.PageSize = value;
            }
        }

        public override IDocumentPaginatorSource Source => _Paginator.Source;

        public DocumentPaginatorWrapper(DocumentPaginator paginator, Size pageSize, Size margin)
        {
            _PageSize = pageSize;
            _Margin = margin;
            _Paginator = paginator;

            _Paginator.PageSize = new Size(_PageSize.Width - margin.Width * 2,
                                            _PageSize.Height - margin.Height * 2);
        }

        Rect Move(Rect rect)
        {
            if (rect.IsEmpty)
            {
                return rect;
            }
            else
            {
                return new Rect(rect.Left + _Margin.Width, rect.Top + _Margin.Height,
                                rect.Width, rect.Height);
            }
        }

        public override DocumentPage GetPage(int pageNumber)
        {
            FlowDocument doc = new FlowDocument();
            DocumentPage page = _Paginator.GetPage(pageNumber);

            // Create a wrapper visual for transformation and add extras
            ContainerVisual newpage = new ContainerVisual();

            DrawingVisual title = new DrawingVisual();

            using (DrawingContext ctx = title.RenderOpen())
            {
                if (_Typeface == null)
                {
                    _Typeface = new Typeface("Times New Roman");
                }

                FormattedText text = new FormattedText("Page " + (pageNumber + 1),
                    System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    _Typeface, 14, Brushes.Black);

                ctx.DrawText(text, new Point(0, -96 / 4)); // 1/4 inch above page content
            }

            DrawingVisual background = new DrawingVisual();

            using (DrawingContext ctx = background.RenderOpen())
            {
                ctx.DrawRectangle(new SolidColorBrush(Color.FromRgb(240, 240, 240)), null, page.ContentBox);
            }

            newpage.Children.Add(background); // Scale down page and center

            ContainerVisual smallerPage = new ContainerVisual();
            smallerPage.Children.Add(page.Visual);
            smallerPage.Transform = new MatrixTransform(0.95, 0, 0, 0.95,
                0.025 * page.ContentBox.Width, 0.025 * page.ContentBox.Height);

            newpage.Children.Add(smallerPage);
            newpage.Children.Add(title);

            newpage.Transform = new TranslateTransform(_Margin.Width, _Margin.Height);

            return new DocumentPage(newpage, _PageSize, Move(page.BleedBox), Move(page.ContentBox));
        } 
        
        

    }
}
