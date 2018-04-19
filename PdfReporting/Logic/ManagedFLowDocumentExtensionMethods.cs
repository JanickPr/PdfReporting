using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;

namespace PdfReporting.Logic
{
    public static class ManagedFLowDocumentExtensionMethods
    {
        public static ManagedFlowDocument InitializeFlowDocumentWith<T>(this ManagedFlowDocument managedFlowDocument, string templateFilePath, T dataSourceObject)
        {
            managedFlowDocument = GetFlowDocumentFrom(templateFilePath);
            managedFlowDocument.SetDataContextTo(dataSourceObject);
            return managedFlowDocument;
            
        }

        private static ManagedFlowDocument GetFlowDocumentFrom(string templateFilePath)
        {
            ManagedFlowDocument managedFlowDocument = new ManagedFlowDocument();
            managedFlowDocument = managedFlowDocument.LoadFromTemplate(templateFilePath);
            return managedFlowDocument;
        }

        public static ManagedFlowDocument LoadFromTemplate(this ManagedFlowDocument managedFlowDocument, string templateFilePath)
        {
            FileInfo templateFile = GetFileInfoFor(templateFilePath);
            using (FileStream fileStream = GetFileStreamFor(templateFile))
            {
                ParserContext parserContext = GetParserContextFor(templateFile);
                managedFlowDocument = (ManagedFlowDocument)XamlReader.Load(fileStream, parserContext);
            }
            return managedFlowDocument;
        }

        private static FileStream GetFileStreamFor(FileInfo file)
        {
            FileInfo fileInfo = GetFileInfoFor(file.FullName);
            return fileInfo.OpenRead();
        }

        private static FileInfo GetFileInfoFor(string filePath)
        {
            return new FileInfo(filePath);
        }

        private static ParserContext GetParserContextFor(FileInfo file)
        {
            Uri uri = GetAbsoluteUriForFile(file);
            ParserContext parser = GetParserContextWithBaseUri(uri);

            return parser;
        }

        private static Uri GetAbsoluteUriForFile(FileInfo file)
        {
            return new Uri(file.FullName, UriKind.Absolute);
        }

        private static ParserContext GetParserContextWithBaseUri(Uri uri)
        {
            return new ParserContext { BaseUri = uri };
        }

        public static void SetDataContextTo<T>(this ManagedFlowDocument managedFlowDocument, T dataSourceObject)
        {
            managedFlowDocument.DataContext = dataSourceObject;
            RenderFLowDocumentWhenDataSourceIsLoaded(managedFlowDocument);
        }

        private static void RenderFLowDocumentWhenDataSourceIsLoaded(ManagedFlowDocument managedFlowDocument)
        {
            managedFlowDocument.Dispatcher.Invoke(DispatcherPriority.Loaded, new Action(() =>
            {
                return;
            }));
        }

        public static Visual GetVisualOfPage(this ManagedFlowDocument managedFlowDocument, int pageIndex)
        {
            DocumentPage page = managedFlowDocument.GetPage(pageIndex);
            return page.Visual;
        }

        public static DocumentPage GetPage(this ManagedFlowDocument managedFlowDocument, int pageIndex)
        {
            DocumentPaginator paginator = managedFlowDocument.GetPaginator();
            return paginator.GetPage(pageIndex);
        }

        public static DocumentPaginator GetPaginator(this ManagedFlowDocument managedFlowDocument)
        {
            return ((IDocumentPaginatorSource)managedFlowDocument).DocumentPaginator;
        }
    }
}
