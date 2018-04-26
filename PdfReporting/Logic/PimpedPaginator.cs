using System.Collections.Generic;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace PdfReporting.Logic
{
    /// <summary>
    /// This paginator provides document headers, footers and repeating table headers 
    /// </summary>
    /// <remarks>
    /// </remarks>
    public class PimpedPaginator : DocumentPaginator
    {
        private readonly DocumentPaginator _basePaginator;
        private readonly ManagedFlowDocument _baseFlowDocument;
        private readonly XpsHeaderAndFooterDefinition _definition;
        private readonly ReportProperties _reportProperties;
        private readonly List<EditablePage> _pages = new List<EditablePage>();
        private bool _isPageCountValid;
        public static int PageCounter;

        public override bool IsPageCountValid
        {
            get => this._isPageCountValid;
        }

        public override int PageCount
        {
            get => this._pages.Count;
        }

        public override Size PageSize
        {
            get => this._basePaginator.PageSize;
            set => this._basePaginator.PageSize = value;
        }

        public override IDocumentPaginatorSource Source
        {
            get => this._basePaginator.Source;
        }

        public PimpedPaginator(ManagedFlowDocument document, XpsHeaderAndFooterDefinition definition, ReportProperties reportProperties)
        {
            this._baseFlowDocument = document.GetCopy();
            this._basePaginator = this._baseFlowDocument.GetPaginator();
            this._definition = definition;
            this._reportProperties = reportProperties;
            this._definition.SetDimensionsOf(this._baseFlowDocument);
        }

        public override DocumentPage GetPage(int pageNumber)
        {
            PageCounter++;

            EditablePage page;
            if(pageNumber < _pages.Count)
                page = _pages[pageNumber];
            else
                page = CreatePageWithContentFrom(pageNumber);
            return new DocumentPage(
                page,
                this._definition.PageSize,
                new Rect(this._definition.PageSize),
                new Rect(this._definition.PageSize)
            );
        }

        private EditablePage CreatePageWithContentFrom(int pageNumber)
        {
            var page = new EditablePage();
            Visual originalPageVisual = this._baseFlowDocument.GetVisualOfPage(pageNumber);
            this._pages.Add(page);
            page.AddVisualAt(this._definition.HeaderHeight, originalPageVisual);
            page.AddFooterAndHeader(this._definition);
            if(this._baseFlowDocument.GetVisualOfPage(pageNumber + 1) == null)
            {
                this._isPageCountValid = true;
                this.AddPageNumbers();
            }

            return page;
        }

        public void AddPageNumbers()
        {
            for(int i = 0; i < this.PageCount; i++)
            {
                EditablePage page = _pages[i];
                if(this._reportProperties.PageNumberSettings.OverallNumeration)
                {
                    page.AddPageNumber(PageCounter, 0, this._reportProperties.PageNumberSettings);
                }
                else
                {
                    page.AddPageNumber(i + 1, this.PageCount, this._reportProperties.PageNumberSettings);
                }
            }
        }
    }
}
