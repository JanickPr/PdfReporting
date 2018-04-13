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
using System.Windows.Threading;

namespace PdfReporting.Logic
{
    public static class ManagedFLowDocumentExtensionMethods
    {
        public static ManagedFlowDocument InitializeFlowDocumentReportWith<T>(this ManagedFlowDocument managedFlowDocument, string templateFilePath, T dataSourceObject)
        {
            managedFlowDocument = GetFlowDocumentFrom(templateFilePath);
            managedFlowDocument.SetUpDataContext(dataSourceObject);
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

        public static void SetUpDataContext<T>(this ManagedFlowDocument managedFlowDocument, T dataSourceObject)
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
    }
}
