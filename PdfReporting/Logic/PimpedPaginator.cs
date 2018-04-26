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
        private int _pageCounter = 0;

        public override bool IsPageCountValid
        {
            get => this._basePaginator.IsPageCountValid;
        }

        public override int PageCount
        {
            get => this._basePaginator.PageCount;
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
            EditablePage page = CreatePageWithContentFrom(pageNumber);
            this._pageCounter++;
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
            page.AddVisualAt(this._definition.HeaderHeight, originalPageVisual);
            page.AddFooterAndHeader(this._definition);
            return AddPageCounterToPage(page, pageNumber);
        }

        private EditablePage AddPageCounterToPage(EditablePage editablePage, int pageNumber)
        {
            if(this._reportProperties.PageNumberSettings.OverallNumeration)
                editablePage.AddPageNumber(this._pageCounter, 0, this._reportProperties.PageNumberSettings);
            else
                editablePage.AddPageNumber(pageNumber, this._basePaginator.PageCount, this._reportProperties.PageNumberSettings);
            return editablePage;
        }
    }
}
