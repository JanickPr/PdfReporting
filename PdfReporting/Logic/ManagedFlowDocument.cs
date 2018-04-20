using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;

namespace PdfReporting.Logic
{
    public class ManagedFlowDocument : FlowDocument
    {
        public new void RemoveLogicalChild(object child)
        {
            base.RemoveLogicalChild(child);
        }

        public static ManagedFlowDocument GetFlowDocumentFrom<T>(string templateFilePath, T dataSourceObject)
        {
            ManagedFlowDocument managedFlowDocument = GetFlowDocumentFrom(templateFilePath);
            managedFlowDocument.SetDataContextTo(dataSourceObject);
            return managedFlowDocument;

        }

        public static ManagedFlowDocument GetFlowDocumentFrom(string templateFilePath)
        {
            FileInfo templateFile = GetFileInfoFor(templateFilePath);
            ManagedFlowDocument managedFlowDocument = CreateManagedFlowDocumentFrom(templateFile);
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

        private static ManagedFlowDocument CreateManagedFlowDocumentFrom(FileInfo templateFile)
        {
            using (FileStream fileStream = GetFileStreamFor(templateFile))
            {
                ParserContext parserContext = GetParserContextFor(templateFile);
                ManagedFlowDocument managedFlowDocument = LoadManagedFlowDocumentFrom(fileStream, parserContext);
                return managedFlowDocument;
            }
        }

        private static ManagedFlowDocument LoadManagedFlowDocumentFrom(FileStream fileStream, ParserContext parserContext)
        {
            return (ManagedFlowDocument)XamlReader.Load(fileStream, parserContext);
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

        public void SetDataContextTo<T>(T dataSourceObject)
        {
            this.DataContext = dataSourceObject;
            RenderFLowDocumentWhenDataSourceIsLoaded(this);
        }

        private void RenderFLowDocumentWhenDataSourceIsLoaded(ManagedFlowDocument managedFlowDocument)
        {
            managedFlowDocument.Dispatcher.Invoke(DispatcherPriority.Loaded, new Action(() =>
            {
                return;
            }));
        }

        public Visual GetVisualOfPage(int pageIndex)
        {
            DocumentPage page = this.GetPage(pageIndex);
            return page.Visual;
        }

        public DocumentPage GetPage(int pageIndex)
        {
            DocumentPaginator paginator = this.GetPaginator();
            return paginator.GetPage(pageIndex);
        }

        public ManagedFlowDocument GetCopy()
        {
            MemoryStream stream = this.SaveToMemoryStream();
            ManagedFlowDocument copy = GetManagedFlowDocumentFrom(stream);
            return copy;
        }

        private MemoryStream SaveToMemoryStream()
        {
            MemoryStream stream = new MemoryStream();
            TextRange textRange = new TextRange(this.ContentStart, this.ContentEnd);
            textRange.Save(stream, DataFormats.Xaml);
            return stream;
        }

        private static ManagedFlowDocument GetManagedFlowDocumentFrom(MemoryStream stream)
        {
            ManagedFlowDocument copy = new ManagedFlowDocument();
            TextRange copyDocumentRange = new TextRange(copy.ContentStart, copy.ContentEnd);
            copyDocumentRange.Load(stream, DataFormats.Xaml);
            return copy;
        }

        public DocumentPaginator GetPaginator()
        {
            return ((IDocumentPaginatorSource)this).DocumentPaginator;
        }
    }
}
