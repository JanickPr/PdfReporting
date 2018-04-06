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
    public static class FLowDocumentExtensionMethods
    {
        public static FlowDocument InitializeFlowDocumentReportWith<T>(this FlowDocument flowDocument, string reportTemplateFilePath, T dataSourceObject)
        {
            flowDocument = GetFlowDocumentFrom(reportTemplateFilePath);
            flowDocument.SetUpDataContext(dataSourceObject);
            return flowDocument;
        }

        private static FlowDocument GetFlowDocumentFrom(string templateFilePath)
        {
            FlowDocument flowDocument = new FlowDocument();
            flowDocument = flowDocument.LoadFromTemplate(templateFilePath);
            return flowDocument;
        }

        public static FlowDocument LoadFromTemplate(this FlowDocument flowDocument, string templateFilePath)
        {
            FileInfo templateFile = GetFileInfoFor(templateFilePath);
            using (FileStream fileStream = GetFileStreamFor(templateFile))
            {
                ParserContext parserContext = GetParserContextFor(templateFile);
                flowDocument = (FlowDocument)XamlReader.Load(fileStream, parserContext);
            }
            return flowDocument;
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

        public static void SetUpDataContext<T>(this FlowDocument flowDocument, T dataSourceObject)
        {
            flowDocument.DataContext = dataSourceObject;
            RenderFLowDocumentWhenDataSourceIsLoaded(flowDocument);
        }

        private static void RenderFLowDocumentWhenDataSourceIsLoaded(FlowDocument flowDocument)
        {
            flowDocument.Dispatcher.Invoke(DispatcherPriority.Loaded, new Action(() =>
            {
                return;
            }));
            //Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
            //{
            //    return;
            //}));
        }
    }
}
