using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.Windows.Xps.Serialization;

namespace PdfReporting.Logic
{
    public static class PdfReportCreation
    {
        #region Fields

        private static double standartPageHeight = 1123.2;
        private static double standartPageWidth = 796.8;
        private static double maxPageHeight;

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
        public static void CreatePdfFromReport<T>(string fileName , T dataSource, string outputDirectory)
        {
            CreatePdfFromReport(fileName, new List<object> { dataSource }, outputDirectory);            
        }







        //public static void SaveAsXps<T>(string fileName, T dataSource)
        //{
        //    object doc;

        //    FileInfo fileInfo = new FileInfo(fileName);

        //    using (FileStream file = fileInfo.OpenRead())
        //    {
        //        ParserContext context = new System.Windows.Markup.ParserContext();
        //        context.BaseUri = new Uri(fileInfo.FullName, UriKind.Absolute);
        //        doc = XamlReader.Load(file, context);
        //    }

        //    ((FlowDocument)doc).DataContext = dataSource;


        //    Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
        //     {

        //         //CreateXpsDocument(doc, fileName);
        //         if (!(doc is IDocumentPaginatorSource))
        //         {
        //             Console.WriteLine("DocumentPaginatorSource expected");
        //         }

        //         using (Package container = Package.Open(fileName + ".xps", FileMode.Create))
        //         {
        //             using (XpsDocument xpsDoc = new XpsDocument(container, CompressionOption.Maximum))
        //             {
        //                 XpsSerializationManager rsm = new XpsSerializationManager(new XpsPackagingPolicy(xpsDoc), false);

        //                 DocumentPaginator paginator = ((IDocumentPaginatorSource)doc).DocumentPaginator;
        //                 paginator.ComputePageCount();
        //                 // 8 inch x 6 inch, with half inch margin
        //                 paginator = new DocumentPaginatorWrapper(paginator, new Size(standartPageWidth, standartPageHeight), new Size(0, 0));

        //                 rsm.SaveAsXaml(paginator);
        //             }
        //         }

        //         Console.WriteLine("{0} generated.", fileName + ".xps");
        //     }));

        //}

        public static void SaveAsXps<T>(string fileName, T dataSource, string outputDirectory)

        {

            object doc;



            FileInfo fileInfo = new FileInfo(fileName);



            using (FileStream file = fileInfo.OpenRead())

            {

                System.Windows.Markup.ParserContext context = new System.Windows.Markup.ParserContext();

                context.BaseUri = new Uri(fileInfo.FullName, UriKind.Absolute);

                doc = System.Windows.Markup.XamlReader.Load(file, context);

            }

            ((FlowDocument)doc).DataContext = dataSource;

           

            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
            {
                Window window = new Window();
                FlowDocumentScrollViewer flowDocumentScrollViewer = new FlowDocumentScrollViewer();
                flowDocumentScrollViewer.Document = (FlowDocument)doc;
                window.Content = flowDocumentScrollViewer;
                window.Show();
                window.Close();

                if (!(doc is IDocumentPaginatorSource))
                {
                    Console.WriteLine("DocumentPaginatorSource expected");
                }

                using (Package container = Package.Open(fileName + ".xps", FileMode.Create))
                {
                    using (XpsDocument xpsDoc = new XpsDocument(container, CompressionOption.Maximum))
                    {
                        XpsSerializationManager rsm = new XpsSerializationManager(new XpsPackagingPolicy(xpsDoc), false);

                        DocumentPaginator paginator = ((IDocumentPaginatorSource)doc).DocumentPaginator;

                        // 8 inch x 6 inch, with half inch margin

                        //paginator = new DocumentPaginatorWrapper(paginator, new Size(standartPageWidth, standartPageHeight), new Size(48, 48));

                        rsm.SaveAsXaml(paginator);
                    }
                }
            }));
        }

        public static void SaveAsXps<T>(string fileName, IEnumerable<T> dataSourceList, string outputDirectory)
        {
            int pageIndexCounter = 0;
            foreach (var dataSource in dataSourceList)
            {


                object doc;

                FileInfo fileInfo = new FileInfo(fileName);

                using (FileStream file = fileInfo.OpenRead())
                {
                    System.Windows.Markup.ParserContext context = new System.Windows.Markup.ParserContext();
                    context.BaseUri = new Uri(fileInfo.FullName, UriKind.Absolute);
                    doc = System.Windows.Markup.XamlReader.Load(file, context);
                }

            ((FlowDocument)doc).DataContext = dataSource;

                Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
                {
                    Window window = new Window();
                    FlowDocumentScrollViewer flowDocumentScrollViewer = new FlowDocumentScrollViewer();
                    flowDocumentScrollViewer.Document = (FlowDocument)doc;
                    window.Content = flowDocumentScrollViewer;
                    window.Show();
                    window.Close();

                    using (MemoryStream memorystream = new MemoryStream())
                    {
                        DocumentPaginator paginator = ((IDocumentPaginatorSource)doc).DocumentPaginator;
                        using (Package container = Package.Open(memorystream, FileMode.Create))
                        {
                            using (XpsDocument xpsDoc = new XpsDocument(container, CompressionOption.Maximum))
                            {
                                XpsSerializationManager rsm = new XpsSerializationManager(new XpsPackagingPolicy(xpsDoc), false);

                                paginator = ((IDocumentPaginatorSource)doc).DocumentPaginator;

                                // 8 inch x 6 inch, with half inch margin

                                //paginator = new DocumentPaginatorWrapper(paginator, new Size(standartPageWidth, standartPageHeight), new Size(48, 48));

                                rsm.SaveAsXaml(paginator);
                                
                            }
                        }

                        CreatePdfFileInDirectory(memorystream, outputDirectory, pageIndexCounter, fileName);

                        pageIndexCounter += paginator.PageCount;
                    }
                }));

            }
        }

        public static XpsDocument CreateXpsDocument(FlowDocument document, string filename)
        {
            using (Package pkg = Package.Open(filename + ".xps", FileMode.Create, FileAccess.ReadWrite))
            {
                string pack = "pack://report.xps";
                PackageStore.RemovePackage(new Uri(pack));
                PackageStore.AddPackage(new Uri(pack), pkg);
                XpsDocument xpsdoc = new XpsDocument(pkg, CompressionOption.NotCompressed, pack);
                XpsSerializationManager rsm = new XpsSerializationManager(new XpsPackagingPolicy(xpsdoc), false);

                DocumentPaginator dp =
                    ((IDocumentPaginatorSource)document).DocumentPaginator;
                rsm.SaveAsXaml(dp);
                return xpsdoc;
            }
        }

        /// <summary>
        /// Sets the  datacontext of the PdfReportBlueprint to the dataSource-object
        /// and creates an .pdf-file in the outputDirectory
        /// </summary>
        /// <param name="pdfReportNamespace">The blueprint for the pdfreport.</param>
        /// <param name="dataSourceList">The datasource which gets binded to the blueprint.</param>
        /// <param name="outputDirectory">The outputdirectory where the .pdf-file is saved to.</param>
        /// <param name="pdfFileName">The name of the .pdf-file that is created.</param>
        public static void CreatePdfFromReport<T>(string pdfReportNamespace, IEnumerable<T> dataSourceList, string outputDirectory)
        {
            using (var stream = System.Reflection.Assembly.GetCallingAssembly().GetManifestResourceStream(pdfReportNamespace))
            {
                System.Windows.Markup.XamlReader.Load(stream);
            }

            
            
            //maxPageHeight = pdfReport.Orientation == Orientation.Vertical ? standartPageHeight : standartPageWidth;

            

            List<FixedPage> pageList = new List<FixedPage>();

            foreach (var dataSource in dataSourceList)
            {

                
            }

            //FixedDocument document = CreateFixedDocumentFromPageList(pageList);
            //MemoryStream xpsDocumentMemoryStream = CreateXpsFromFixedDocument(document);
            //CreatePdfFileInDirectory(xpsDocumentMemoryStream, outputDirectory);

        }

        private static void CreatePdfFileInDirectory(MemoryStream xpsDocumentMemoryStream, string outputDirectory, int pageIndex, string filename)
        {
            if (outputDirectory == null)
            {
                return;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(outputDirectory));

            bool fileCreated = false;

            using (PdfSharp.Xps.XpsModel.XpsDocument pdfXpsDoc = PdfSharp.Xps.XpsModel.XpsDocument.Open(xpsDocumentMemoryStream))
            {
                if(!File.Exists(outputDirectory))
                {
                    PdfSharp.Xps.XpsConverter.Convert(pdfXpsDoc, outputDirectory, 0);
                    fileCreated = true;
                }
                else
                {
                    PdfSharp.Xps.XpsConverter.Convert(pdfXpsDoc, filename.Replace(".xaml", "") + ".pdf", 0);
                }
            }
            if (fileCreated) return; 

            var pdfDokument = new PdfSharp.Pdf.PdfDocument(outputDirectory);
            var tempPdfDocument = new PdfSharp.Pdf.PdfDocument(filename + ".pdf");

            for (int i = 0; i < tempPdfDocument.PageCount; i++)
            {
                pdfDokument.AddPage(tempPdfDocument.Pages[i]);
            }

            pdfDokument.Dispose();
            tempPdfDocument.Dispose();


            //Hier wird die leere erste Seite wieder entfernt, die in ErstellePdfInhalt<T> erstellt wurde.
            //var pdfDokument = new PdfSharp.Pdf.PdfDocument(outputDirectory);
            //pdfDokument.Pages.RemoveAt(0);
            //pdfDokument.Save(outputDirectory);
        }

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="element"></param>
        ///// <returns></returns>
        //private static FrameworkElement GetChild(FrameworkElement element, int childIndex)
        //{
        //    if (IsSameOrSubclass(typeof(Grid), element.GetType()))
        //    {
        //        if (((Grid)element).Children.Count <= childIndex)
        //            throw new IndexOutOfRangeException();

        //        return (FrameworkElement)((Grid)element).Children[childIndex];
        //    }
        //    else if(IsSameOrSubclass(typeof(Panel), element.GetType()))
        //    {
        //        if (((Panel)element).Children.Count <= childIndex)
        //            throw new IndexOutOfRangeException();

        //        return (FrameworkElement)((Panel)element).Children[childIndex];
        //    }
        //    else if(IsSameOrSubclass(typeof(ContentControl), element.GetType()))
        //    {
        //        return (FrameworkElement)((ContentControl)element).Content;
        //    }
        //    else if (IsSameOrSubclass(typeof(Decorator), element.GetType()))
        //    {
        //        return (FrameworkElement)((Decorator)element).Child;
        //    }

        //    throw new ArgumentException();
        //}

        //private static MemoryStream CreateXpsFromFixedDocument(FixedDocument document)
        //{
        //    MemoryStream memoryStream = new MemoryStream();
        //    var package = Package.Open(memoryStream, FileMode.Create);
        //    var xpsDoc = new XpsDocument(package, CompressionOption.Normal);
        //    XpsDocumentWriter xpsWriter = XpsDocument.CreateXpsDocumentWriter(xpsDoc);

        //    // xps documents are built using fixed document sequences
        //    var fixedDocSeq = new FixedDocumentSequence();
        //    var docRef = new DocumentReference();
        //    docRef.BeginInit();
        //    docRef.SetDocument(document);
        //    docRef.EndInit();
        //    ((IAddChild)fixedDocSeq).AddChild(docRef);

        //    // write out our fixed document to xps
        //    xpsWriter.Write(fixedDocSeq.DocumentPaginator);

        //    xpsDoc.Close();
        //    package.Close();

        //    return memoryStream;
        //}

        ///// <summary>
        ///// Creates a FixedDocument which contains every FixedPage from the given list.
        ///// </summary>
        ///// <param name="pageList"></param>
        ///// <returns></returns>
        //private static FixedDocument CreateFixedDocumentFromPageList(List<FixedPage> pageList)
        //{
        //    FixedDocument document = new FixedDocument();

        //    pageList.ForEach(page =>
        //    {
        //        PageContent pageContent = new PageContent();

        //        ((IAddChild)pageContent).AddChild(page);
        //        document.Pages.Add(pageContent);
        //    });

        //    return document;
        //}

        ///// <summary>
        ///// Sets the datacontext for the pdfReportBlueprint.
        ///// </summary>
        ///// <param name="pdfReport"></param>
        ///// <param name="dataSource"></param>
        ///// <returns></returns>
        //private static PdfReport SetDataContext(PdfReport pdfReport, object dataSource)
        //{
        //    pdfReport.DataContext = dataSource;
        //    return pdfReport;
        //}

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="pageTemplate"></param>
        ///// <param name="sourceElement"></param>
        ///// <param name="iterationLevel"></param>
        ///// <param name="pageList"></param>
        ///// <returns></returns>
        //private static List<FixedPage> DivideElementIntoPages
        //    (PdfReport pageTemplate, FrameworkElement sourceElement, int iterationLevel = 0, List<FixedPage> pageList = null)
        //{
        //    if (pageList == null)
        //        pageList = new List<FixedPage> { GenerateNewFixedPage(pageTemplate) };


        //    // If the sourceElement fits on the page, than it 
        //    if (ElementFitsOnPage(sourceElement, pageList.Last()))
        //    {
        //        FixedPage lastPage = pageList.Last();
        //        AddElementToPage(sourceElement, ref lastPage, iterationLevel);
        //        pageList[pageList.Count - 1] = lastPage;
        //        return pageList;
        //    }
            
        //    // If the sourceElement is on the Lowest level, has no children and does not fit on the last page
        //    // a new page is created with the sourceElement and the complete structure of it.
        //    if(VisualTreeHelper.GetChildrenCount(sourceElement) < 1)
        //    {
        //        FixedPage page = GenerateNewFixedPage(pageTemplate);
        //        AddElementToNewPage(sourceElement, ref page);
        //        pageList.Add(page);
        //        return pageList;
        //    }

        //    int childrenCount = VisualTreeHelper.GetChildrenCount(sourceElement);
        //    FrameworkElement currentChild = new FrameworkElement(); 

        //    for (int i = 0; i < childrenCount; i++)
        //    {
        //        currentChild = (FrameworkElement)VisualTreeHelper.GetChild(sourceElement, i);

        //        pageList = DivideElementIntoPages(pageTemplate, currentChild, iterationLevel + 1, pageList);
        //    }

        //    return pageList;
        //}

        //private static FrameworkElement GetLowestChildElement(FrameworkElement element)
        //{
        //    while(VisualTreeHelper.GetChildrenCount(element) > 0)
        //    {
        //        int childrenCount = VisualTreeHelper.GetChildrenCount(element);
        //        element = (FrameworkElement)VisualTreeHelper.GetChild(element, childrenCount - 1);
        //    }

        //    return element;
        //}

        ///// <summary>
        ///// Adds the complete structure that houses the given element to the given Page.
        ///// </summary>
        ///// <param name="element"></param>
        ///// <param name="page"></param>
        //private static void AddElementToNewPage(FrameworkElement element, ref FixedPage page)
        //{
        //    List<FrameworkElement> elementParents = GetAllParentElements(element);

        //    for (int i = 0; i < elementParents.Count; i++)
        //    {
        //        FrameworkElement parentElement = elementParents[i];
        //        FrameworkElement childElement = elementParents[i + 1];

        //        if (parentElement == elementParents.Last())
        //            childElement = element;

        //        if (i == 0)
        //            page.Children.Add(parentElement);

               
        //    }
        //}

        ///// <summary>
        ///// Returns all frameworkelements that are in a parental relationship to the given element.
        ///// </summary>
        ///// <param name="element"></param>
        ///// <returns></returns>
        //private static List<FrameworkElement> GetAllParentElements(FrameworkElement element)
        //{
        //    List<FrameworkElement> elementParents = new List<FrameworkElement>();

        //    FrameworkElement currentElement = element;

        //    while (VisualTreeHelper.GetParent(currentElement) != null)
        //    {
        //        currentElement = (FrameworkElement)VisualTreeHelper.GetParent(currentElement);
        //        currentElement = RemoveContentFromElement(currentElement);
        //        elementParents.Insert(0, currentElement);
        //    }

        //    return elementParents;
        //}

        ///// <summary>
        ///// Evaluates if the given element fits on the given page.
        ///// </summary>
        ///// <param name="element"></param>
        ///// <returns></returns>
        //private static bool ElementFitsOnPage(FrameworkElement element, FixedPage page)
        //{
        //    double pageHeight = 0;
        //    double elementHeight = element.ActualHeight;
        //    double pageChildrenCount = page.Children.Count;
            
        //    for (int i = 0; i < pageChildrenCount; i++)
        //    {
        //        FrameworkElement childElement = (FrameworkElement)page.Children[i];
        //        pageHeight += childElement.ActualHeight;
        //    }

        //    return pageHeight + elementHeight <= maxPageHeight;
        //}

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <returns></returns>
        //private static FixedPage GenerateNewFixedPage(PdfReport pageTemplate)
        //{
        //    FixedPage page = new FixedPage();

        //    if(pageTemplate.Header != null)
        //    {
        //        page.Children.Add(pageTemplate.Header);
        //    }
        //    return page;
        //}

        ///// <summary>
        ///// Returns the given FrameworkElement without all its content.
        ///// </summary>
        ///// <param name="frameworkElement"></param>
        ///// <returns></returns>
        //private static FrameworkElement RemoveContentFromElement(FrameworkElement frameworkElement)
        //{
        //    if(frameworkElement.GetType() == typeof(ContentControl))
        //    {
        //        ((ContentControl)frameworkElement).Content = null;
        //    }
        //    else if(frameworkElement.GetType() == typeof(Panel))
        //    {
        //        ((Panel)frameworkElement).Children.Clear();
        //    }

        //    return frameworkElement;
        //}

        ///// <summary>
        ///// Adds the given FrameworkElement to the last element on the given treeLevel of the given FixedPage-object.
        ///// </summary>
        ///// <param name="elementToBeAdded"></param>
        ///// <param name="page"></param>
        ///// <param name="treeLevel"></param>
        //private static void AddElementToPage(FrameworkElement elementToBeAdded, ref FixedPage page, int treeLevel)
        //{
        //    FrameworkElement currentElement = (FrameworkElement)page;
        //    while(treeLevel > 0)
        //    {
        //        if (currentElement.GetType() == typeof(ContentControl))
        //            currentElement = (FrameworkElement) ((ContentControl)currentElement).Content;
        //        else if (currentElement.GetType() == typeof(Panel))
        //            currentElement = (FrameworkElement)((Panel)currentElement).Children[((Panel)currentElement).Children.Count];

        //        treeLevel--;
        //    }

        //    AddElement(ref currentElement, elementToBeAdded);
        //}

        ///// <summary>
        ///// Adds the given elementToBeAdded to the parentElement.
        ///// </summary>
        ///// <param name="parentElement"></param>
        ///// <param name="elementToBeAdded"></param>
        //private static void AddElement(ref FrameworkElement parentElement, FrameworkElement elementToBeAdded)
        //{
        //    SeperateFromParent(elementToBeAdded);
        //    if (parentElement.GetType() == typeof(ContentControl) && ((ContentControl)parentElement).Content == null)
        //    {
        //        ((ContentControl)parentElement).Content = elementToBeAdded;
        //    }

        //    else if (parentElement.GetType() == typeof(Panel))
        //    {
        //        ((Panel)parentElement).Children.Add(elementToBeAdded);
        //    }

        //    else if (parentElement.GetType() == typeof(FixedPage))
        //    {
        //        ((FixedPage)parentElement).Children.Add(elementToBeAdded);
        //    }

        //    else
        //        throw new Exception("Error while adding VisualTreeElement to page.");
        //}


        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="panel"></param>
        ///// <returns></returns>
        //private static bool PanelHasVerticalOrientedItems(Panel panel)
        //{
        //    if (panel.GetType() == typeof(Grid))
        //    {
        //        Grid grid = (Grid)panel;
        //        return grid.RowDefinitions.Count > 0;
        //    }

        //    return panel.LogicalOrientationPublic == Orientation.Vertical;

        //    //Here we must check for other paneltypes e.g. Canvas or WrapPanel
        //}
        
        ///// <summary>
        ///// Returns the height that is used by the given frameworkelement with the content included.
        ///// </summary>
        ///// <param name="frameworkElement"></param>
        ///// <returns></returns>
        //private static double GetFrameworkElementHeightIncludingContent(FrameworkElement frameworkElement)
        //{
        //    double height = GetFrameworkElementHeightExcludingContent(frameworkElement);

        //    height += frameworkElement.ActualHeight;

        //    return height;
        //}

        ///// <summary>
        ///// Returns the height that is used by the given frameworkelement without the content.
        ///// </summary>
        ///// <param name="frameworkElement"></param>
        ///// <returns></returns>
        //private static double GetFrameworkElementHeightExcludingContent(FrameworkElement frameworkElement)
        //{
        //    double height = 0;

        //    height += GetHeightFromSpacingproperty(frameworkElement.Margin);

        //    if (frameworkElement.GetType().GetProperty("Padding") != null)
        //    {
        //        var padding = frameworkElement.GetType().GetProperty("Padding").GetValue(frameworkElement);
        //        height += GetHeightFromSpacingproperty((Thickness)padding);
        //    }

        //    if (frameworkElement.GetType().GetProperty("BorderThickness") != null)
        //    {
        //        var borderThickness = frameworkElement.GetType().GetProperty("BorderThickness").GetValue(frameworkElement);
        //        height += GetHeightFromSpacingproperty((Thickness)borderThickness);
        //    }

        //    return height;
        //}

        //private static void SeperateFromParent(FrameworkElement element)
        //{
        //    var parentElement = element.Parent;

        //    if (parentElement.GetType() == typeof(ContentControl) && ((ContentControl)parentElement).Content == null)
        //    {
        //        ((ContentControl)parentElement).Content = null;
        //    }

        //    else if (parentElement.GetType() == typeof(Panel))
        //    {
        //        ((Panel)parentElement).Children.Remove(element);
        //    }

        //    else if (parentElement.GetType() == typeof(FixedPage))
        //    {
        //        ((FixedPage)parentElement).Children.Remove(element);
        //    }

        //    else
        //        throw new Exception("Error while adding VisualTreeElement to page.");
        //}

        ///// <summary>
        ///// returns the vertical space that is used for spacing such as margin or padding.
        ///// </summary>
        ///// <param name="spacing"></param>
        ///// <returns></returns>
        //private static double GetHeightFromSpacingproperty(Thickness spacing)
        //{
        //    return spacing.Top + spacing.Bottom;
        //}

        ///// <summary>
        ///// Returns true if the given descendant type derives from the given base type
        ///// or if the decendant type is equal to the base type.
        ///// </summary>
        ///// <param name="Base"></param>
        ///// <param name="Descendant"></param>
        ///// <returns></returns>
        //private static bool IsSameOrSubclass(Type Base, Type Descendant)
        //{
        //    return Descendant.IsSubclassOf(Base)
        //           || Descendant == Base;
        //}

        #endregion
    }
}
