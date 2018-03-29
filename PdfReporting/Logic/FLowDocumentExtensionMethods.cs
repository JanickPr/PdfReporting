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

namespace PdfReporting.Logic
{
    public static class FLowDocumentExtensionMethods
    {
        public static void LoadFromTemplate(this FlowDocument flowDocument, string templateFileName)
        {
            using (FileStream fileStream = GetFileStreamFor(templateFileName))
            {
                ParserContext parserContext = GetParserContextFor(templateFileName);
                flowDocument = (FlowDocument)XamlReader.Load(fileStream, parserContext);
            }
        }

        private static ParserContext GetParserContextFor(string filename)
        {
            Uri uri = GetAbsoluteUriForFile(filename);
            ParserContext parser = GetParserContextWithBaseUri(uri);

            return parser;
        }

        private static Uri GetAbsoluteUriForFile(string fullFileName)
        {
            return new Uri(fullFileName, UriKind.Absolute);
        }

        private static ParserContext GetParserContextWithBaseUri(Uri uri)
        {
            return new ParserContext { BaseUri = uri };
        }

        private static FileStream GetFileStreamFor(string fileName)
        {
            FileInfo fileInfo = GetFileInfoFor(fileName);
            return fileInfo.OpenRead();
        }

        private static FileInfo GetFileInfoFor(string filename)
        {
            return new FileInfo(filename);
        }

        public static void SetUpDataContext<T>(this FlowDocument flowDocument, T dataSourceObject)
        {
            flowDocument.DataContext = dataSourceObject;
            RenderFLowDocumentWhenDataSourceIsLoaded(flowDocument);
        }

        private static void RenderFLowDocumentWhenDataSourceIsLoaded(FlowDocument flowDocument)
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
            {
                RenderFlowDocument(flowDocument);
            }));
        }

        private static void RenderFlowDocument(FlowDocument flowDocument)
        {
            FlowDocumentScrollViewer flowDocumentScrollViewer = new FlowDocumentScrollViewer { Document = flowDocument };
            Window window = new Window { Content = flowDocumentScrollViewer };
            window.Show();
            window.Close();
        }

        private
    }
}
