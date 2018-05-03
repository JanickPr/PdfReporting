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
        private readonly ReportContentDefinition _definition;
        private readonly ReportProperties _reportProperties;
        private readonly List<EditablePage> _pages = new List<EditablePage>();
        private int _pageCount = -1;
        public static int GlobalPageCounter;

        public override bool IsPageCountValid
        {
            get => this._pages.Count == this.PageCount;
        }

        public override int PageCount
        {
            get
            {
                if(this._pageCount == -1)
                    this._pageCount = GetPageCount();
                return this._pageCount;
            }
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

        public PimpedPaginator(ManagedFlowDocument document, ReportContentDefinition definition, ReportProperties reportProperties)
        {
            this._baseFlowDocument = document;//document.GetCopy();
            this._basePaginator = this._baseFlowDocument.GetPaginator();
            this._definition = definition;
            this._reportProperties = reportProperties;
            this._definition.SetDimensionsOf(this._baseFlowDocument);
        }

        public override DocumentPage GetPage(int pageNumber)
        {
            GlobalPageCounter++;

            EditablePage page;
            if(pageNumber < this._pages.Count)
                page = this._pages[pageNumber];
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

            this.AddPageNumber(pageNumber);
            return page;
        }

        private int GetPageCount()
        {
            int i = 0;
            while(this._baseFlowDocument.GetVisualOfPage(i) != null)
            {
                i++;
            }
            return i;
        }

        public void AddPageNumber(int pageNumber)
        {
            EditablePage page = this._pages[pageNumber];
            if(this._reportProperties.PageNumberSettings.OverallNumeration)
            {
                page.AddPageNumber(GlobalPageCounter, 0, this._reportProperties.PageNumberSettings);
            }
            else
            {
                page.AddPageNumber(pageNumber + 1, this.PageCount, this._reportProperties.PageNumberSettings);
            }
        }
    }
}
