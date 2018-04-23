using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Controls;
using System.Windows.Media;
using System.Globalization;
using System.IO;
using PdfReporting.Logic;

namespace PdfReporting.Logic
{
	
	/// <summary>
	/// This paginator provides document headers, footers and repeating table headers 
	/// </summary>
	/// <remarks>
	/// </remarks>
	public class PimpedPaginator : DocumentPaginator {

        private ContainerVisual _currentHeader = null;
        private DocumentPaginator _basePaginator;
        private ManagedFlowDocument _baseFlowDocument;
        private XpsHeaderAndFooterDefinition _definition;
        private ReportProperties _reportProperties;
        private int _pageCounter = 0;

        public override bool IsPageCountValid => _basePaginator.IsPageCountValid;

        public override int PageCount => _basePaginator.PageCount;

        public override Size PageSize
        {
            get => _basePaginator.PageSize;
            set => _basePaginator.PageSize = value;
        }

        public override IDocumentPaginatorSource Source => _basePaginator.Source;

        public PimpedPaginator(ManagedFlowDocument document, XpsHeaderAndFooterDefinition definition, ReportProperties reportProperties)
        {
            this._baseFlowDocument = document.GetCopy();
			this._basePaginator = _baseFlowDocument.GetPaginator();
			this._definition = definition;
            this._reportProperties = reportProperties;
            SetDimensionsOf(_baseFlowDocument);
        }

        private ManagedFlowDocument SetDimensionsOf(ManagedFlowDocument managedFlowDocument)
        {
            managedFlowDocument.ColumnWidth = double.MaxValue; // Prevent columns
            managedFlowDocument.PageWidth = _definition.ContentSize.Width;
            managedFlowDocument.PageHeight = _definition.ContentHeight;
            managedFlowDocument.PagePadding = new Thickness(0);
            return managedFlowDocument;
        }

		public override DocumentPage GetPage(int pageNumber)
        {
            EditablePage page = CreatePageWithContentFrom(pageNumber);
            _pageCounter++;
            return new DocumentPage(
				page, 
				_definition.PageSize, 
				new Rect(_definition.PageSize),
				new Rect(_definition.ContentSize)
			);
		}

        private EditablePage CreatePageWithContentFrom(int pageNumber)
        {
            EditablePage page = CreatePageWithHeaderFooter();
            Visual originalPageVisual = _baseFlowDocument.GetVisualOfPage(pageNumber);
            page.AddVisualAt(_definition.HeaderHeight, originalPageVisual);
            page = AddPageCounterToPage(page, pageNumber);
            return page;
        }

        private EditablePage CreatePageWithHeaderFooter()
        {
            EditablePage page = new EditablePage();
            page.AddVisualAtTopOfPage(_definition.HeaderVisual);
            page.AddVisualAt(_definition.FooterOffsetY, _definition.FooterVisual);
            return page;
        }

        private EditablePage AddPageCounterToPage(EditablePage editablePage, int pageNumber)
        {
            if (_reportProperties.PageNumberSettings.OverallNumeration)
                editablePage.AddPageNumber(_pageCounter, default, _reportProperties.PageNumberSettings);
            else
                editablePage.AddPageNumber(pageNumber, _basePaginator.PageCount, _reportProperties.PageNumberSettings);
            return editablePage;
        }
	}
}
