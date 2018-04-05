using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Interop;
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

        public static async void CreatePdfReportFromObjectAsync<T>(string templateFilePath, T dataSourceObject, string outputDirectory, IProgress<int> progress = null)
            => await CreatePdfReportFromObjectListAsync(templateFilePath, new List<T> { dataSourceObject }, outputDirectory, progress);

        public static async Task CreatePdfReportFromObjectListAsync<T>(string templateFilePath, IEnumerable<T> dataSourceList, string outputDirectory, IProgress<int> progress = null)
        {
            var task = Task.Factory.StartNew(() =>
            {
                XpsDocumentSplicer xpsDocumentSplicer = new XpsDocumentSplicer();

                foreach (var dataSourceItem in dataSourceList)
                {
                    progress.Report(GetProcessingProgress(dataSourceItem, dataSourceList));
                    xpsDocumentSplicer.AddXpsDocumentFrom(templateFilePath, dataSourceItem);
                }

                SaveAsPdf(xpsDocumentSplicer, outputDirectory);
            });

            await task;
        }

        private static int GetProcessingProgress<T>(T dataSourceItem, IEnumerable<T> dataSourceList)
        {
            List<T> tempList = dataSourceList.ToList();
            int index = tempList.IndexOf(dataSourceItem);
            return (index / dataSourceList.Count()) * 100;
        }

        private static void SaveAsPdf(XpsDocumentSplicer xpsDocumentSplicer, string outputDirectory)
        {
            string tempXpsSavePath = GetTempXpsSavePath();
            xpsDocumentSplicer.SaveSplicedXpsDocumentTo(tempXpsSavePath);
            ConvertXpsToPdf(tempXpsSavePath, outputDirectory);
            File.Delete(tempXpsSavePath);
        }

        private static string GetTempXpsSavePath()
        {
            return AppDomain.CurrentDomain.BaseDirectory + "\tempXpsFile.xps";
        }

        private static void ConvertXpsToPdf(string xpsFileName, string outputDirectory)
        {
            using (PdfSharp.Xps.XpsModel.XpsDocument pdfXpsDoc = PdfSharp.Xps.XpsModel.XpsDocument.Open(xpsFileName))
            {
                PdfSharp.Xps.XpsConverter.Convert(pdfXpsDoc, outputDirectory, 0);
            }
        }

        
        #endregion
    }
}
