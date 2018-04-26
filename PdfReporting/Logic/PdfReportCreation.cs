using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace PdfReporting.Logic
{
    public static class PdfReportCreation
    {
        #region Methods

        public static async Task CreatePdfReportFromObjectAsync<T>(ReportProperties reportProperties, T dataSourceObject, CancellationToken token = default, IProgress<int> progress = null)
        {
            await CreatePdfReportFromObjectListAsync(reportProperties, new List<T> { dataSourceObject }, token, progress).ConfigureAwait(false);
        }

        public static Task CreatePdfReportFromObjectListAsync<T>(ReportProperties reportProperties, IEnumerable<T> dataSourceObjectList, CancellationToken token = default, IProgress<int> progress = null)
        {
            var taskCompletionSource = new TaskCompletionSource<object>();
            var thread = new Thread(() => CreatePdfReportFromObjectList(reportProperties, dataSourceObjectList, token, progress));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return taskCompletionSource.Task;
        }

        private static void CreatePdfReportFromObjectList<T>(ReportProperties reportProperties, IEnumerable<T> dataSourceObjectList, CancellationToken token = default, IProgress<int> progress = null)
        {
            if(dataSourceObjectList == null)
                throw new ArgumentNullException(nameof(dataSourceObjectList));

            var xpsDocumentSplicer = new XpsDocumentSplicer(reportProperties);
            foreach(T dataSourceItem in dataSourceObjectList)
            {
                progress?.Report(GetProcessingProgress(dataSourceItem, dataSourceObjectList));
                if(token.IsCancellationRequested)
                    return;
                xpsDocumentSplicer.AddXpsDocumentWith(dataSourceItem);
            }
            SaveAsPdf(xpsDocumentSplicer, reportProperties.OutputDirectory);
            PimpedPaginator.PageCounter = 0;
        }

        private static int GetProcessingProgress<T>(T dataSourceItem, IEnumerable<T> dataSourceList)
        {
            var tempList = dataSourceList.ToList();
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
            return AppDomain.CurrentDomain.BaseDirectory + "tempXpsFile.xps";
        }

        private static void ConvertXpsToPdf(string xpsFileName, string outputDirectory)
        {
            using(var pdfXpsDoc = PdfSharp.Xps.XpsModel.XpsDocument.Open(xpsFileName))
            {
                PdfSharp.Xps.XpsConverter.Convert(pdfXpsDoc, outputDirectory, 0);
            }
        }

        #endregion
    }
}
