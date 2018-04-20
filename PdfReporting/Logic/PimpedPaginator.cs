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

        ContainerVisual _currentHeader = null;
        private DocumentPaginator _paginator;
        private XpsHeaderAndFooterDefinition _definition;

        public override bool IsPageCountValid => _paginator.IsPageCountValid;
        
        public override int PageCount => _paginator.PageCount;

        public override IDocumentPaginatorSource Source => _paginator.Source;

        public override Size PageSize
        {
            get{ return _paginator.PageSize; }
            set{ _paginator.PageSize = value; }
        }

        


        public PimpedPaginator(ManagedFlowDocument document, XpsHeaderAndFooterDefinition def)
        {
            ManagedFlowDocument copy = document.GetCopy();
			this._paginator = copy.GetPaginator();
			this._definition = def;
			_paginator.PageSize = new Size(_definition.ContentSize.Width, 500);
            copy = this.SetDimensionsOf(copy);

        }

        private ManagedFlowDocument SetDimensionsOf(ManagedFlowDocument managedFlowDocument)
        {
            managedFlowDocument.ColumnWidth = double.MaxValue; // Prevent columns
            managedFlowDocument.PageWidth = _definition.ContentSize.Width;
            managedFlowDocument.PageHeight = _definition.ContentHeight;
            managedFlowDocument.PagePadding = new Thickness(0);
            return managedFlowDocument;
        }

		public override DocumentPage GetPage(int pageNumber) {
			// Use default paginator to handle pagination
			Visual originalPage = _paginator.GetPage(pageNumber).Visual;

			//Console.WriteLine("--- Begin Page {0} -------", pageNumber + 1);
			//originalPage.DumpVisualTree(Console.Out); //die Methode existiert nicht
			//Console.WriteLine("--- End Page {0} -------", pageNumber + 1);
			//Console.WriteLine();

			ContainerVisual pageVisual = new ContainerVisual();
            ContainerVisual pageContentVisual = new ContainerVisual()
            {
                Transform = new TranslateTransform(
                    0,
                    _definition.HeaderHeight
                )
            };
            pageContentVisual.Children.Add(originalPage);
            pageVisual.Children.Add(pageContentVisual);

            if (_definition.HeaderVisual != null)
            {
                ContainerVisual pageHeaderVisual = new ContainerVisual()
                {
                    Transform = new TranslateTransform(0, 0)
                };
                pageHeaderVisual.Children.Add(_definition.HeaderVisual);
                pageVisual.Children.Add(pageHeaderVisual);
            }

            if (_definition.FooterVisual != null)
            {
                ContainerVisual pageFooterVisual = new ContainerVisual()
                {
                    Transform = new TranslateTransform(0, _definition.PageSize.Height - _definition.FooterHeight),
                };
                pageFooterVisual.Children.Add(_definition.FooterVisual);
                pageVisual.Children.Add(pageFooterVisual);
            }




            //Create headers and footers



            //Check for repeating table headers
            if (_definition.RepeatTableHeaders)
            {
                // Find table header
                ContainerVisual table;
                if (PageStartsWithTable(originalPage, out table) && _currentHeader != null)
                {
                    // The page starts with a table and a table header was
                    // found on the previous page. Presumably this table 
                    // was started on the previous page, so we'll repeat the
                    // table header.
                    Rect headerBounds = VisualTreeHelper.GetDescendantBounds(_currentHeader);
                    Vector offset = VisualTreeHelper.GetOffset(_currentHeader);
                    ContainerVisual tableHeaderVisual = new ContainerVisual();

                    // Translate the header to be at the top of the page
                    // instead of its previous position
                    tableHeaderVisual.Transform = new TranslateTransform(
                        _definition.ContentOrigin.X,
                        _definition.ContentOrigin.Y - headerBounds.Top
                    );

                    // Since we've placed the repeated table header on top of the
                    // content area, we'll need to scale down the rest of the content
                    // to accomodate this. Since the table header is relatively small,
                    // this probably is barely noticeable.
                    double yScale = (_definition.ContentSize.Height - headerBounds.Height) / _definition.ContentSize.Height;
                    TransformGroup group = new TransformGroup();
                    group.Children.Add(new ScaleTransform(1.0, yScale));
                    group.Children.Add(new TranslateTransform(
                        _definition.ContentOrigin.X,
                        _definition.ContentOrigin.Y + headerBounds.Height
                    ));
                    pageContentVisual.Transform = group;

                    ContainerVisual cp = VisualTreeHelper.GetParent(_currentHeader) as ContainerVisual;
                    if (cp != null)
                    {
                        cp.Children.Remove(_currentHeader);
                    }
                    tableHeaderVisual.Children.Add(_currentHeader);
                    pageVisual.Children.Add(tableHeaderVisual);
                }

                // Check if there is a table on the bottom of the page.
                // If it's there, its header should be repeated
                ContainerVisual newTable, newHeader;
                if (PageEndsWithTable(originalPage, out newTable, out newHeader))
                {
                    if (newTable == table)
                    {
                        // Still the same table so don't change the repeating header
                    }
                    else
                    {
                        // We've found a new table. Repeat the header on the next page
                        _currentHeader = newHeader;
                    }
                }
                else
                {
                    // There was no table at the end of the page
                    _currentHeader = null;
                }
            }

            return new DocumentPage(
				pageVisual, 
				_definition.PageSize, 
				new Rect(_definition.PageSize),
				new Rect(_definition.ContentSize)
			);
		}



        /// <summary>
        /// Checks if the page ends with a table.
        /// </summary>
        /// <remarks>
        /// There is no such thing as a 'TableVisual'. There is a RowVisual, which
        /// is contained in a ParagraphVisual if it's part of a table. For our
        /// purposes, we'll consider this the table Visual
        /// 
        /// You'd think that if the last element on the page was a table row, 
        /// this would also be the last element in the visual tree, but this is not true
        /// The page ends with a ContainerVisual which is aparrently  empty.
        /// Therefore, this method will only check the last child of an element
        /// unless this is a ContainerVisual
        /// </remarks>
        /// <param name="originalPage"></param>
        /// <returns></returns>
        private bool PageEndsWithTable(DependencyObject element, out ContainerVisual tableVisual, out ContainerVisual headerVisual) {
			tableVisual = null;
			headerVisual = null;
			if(element.GetType().Name == "RowVisual") {
				tableVisual = (ContainerVisual)VisualTreeHelper.GetParent(element);
				headerVisual = (ContainerVisual)VisualTreeHelper.GetChild(tableVisual, 0);
				return true;
			}
			int children = VisualTreeHelper.GetChildrenCount(element);
			if(element.GetType() == typeof(ContainerVisual)) {
				for(int c = children - 1; c >= 0; c--) {
					DependencyObject child = VisualTreeHelper.GetChild(element, c);
					if(PageEndsWithTable(child, out tableVisual, out headerVisual)) {
						return true;
					}
				}
			} else if(children > 0) {
				DependencyObject child = VisualTreeHelper.GetChild(element, children - 1);
				if(PageEndsWithTable(child, out tableVisual, out headerVisual)) {
					return true;
				}
			}
			return false;
		}


		/// <summary>
		/// Checks if the page starts with a table which presumably has wrapped
		/// from the previous page.
		/// </summary>
		/// <param name="element"></param>
		/// <param name="tableVisual"></param>
		/// <param name="headerVisual"></param>
		/// <returns></returns>
		private bool PageStartsWithTable(DependencyObject element, out ContainerVisual tableVisual) {
			tableVisual = null;
			if(element.GetType().Name == "RowVisual") {
				tableVisual = (ContainerVisual)VisualTreeHelper.GetParent(element);
				return true;
			}
			if(VisualTreeHelper.GetChildrenCount(element)> 0) {
				DependencyObject child = VisualTreeHelper.GetChild(element, 0);
				if(PageStartsWithTable(child, out tableVisual)) {
					return true;
				}
			}
			return false;
		}


		#region DocumentPaginator members

		

		#endregion

	}
}
