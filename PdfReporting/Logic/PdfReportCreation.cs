using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Amib.Threading;

namespace PdfReporting.Logic
{
    public static class PdfReportCreation
    {
        private static double _progress;
        private static CancellationToken _cancellationToken;
        private static IProgress<double> _progressReporter;

        #region Methods

        public static async Task CreatePdfReportFromObjectAsync<T>(ReportProperties reportProperties, T dataSourceObject, CancellationToken token = default, IProgress<double> progress = null)
        {
            await CreatePdfReportFromObjectListAsync(reportProperties, new List<T> { dataSourceObject }, token, progress).ConfigureAwait(false);
        }

        public static Task CreatePdfReportFromObjectListAsync<T>(ReportProperties reportProperties, IEnumerable<T> dataSourceObjectList, CancellationToken token = default, IProgress<double> progress = null)
        {
            var taskCompletionSource = new TaskCompletionSource<object>();
            var thread = new Thread(() => CreatePdfReportFromObjectList(reportProperties, dataSourceObjectList, token, progress));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return taskCompletionSource.Task;
        }

        public static void CreatePdfReportFromObjectList<T>(ReportProperties reportProperties, IEnumerable<T> dataSourceObjectList, CancellationToken cancellationToken = default, IProgress<double> progressReporter = null)
        {
            if(dataSourceObjectList == null)
                throw new ArgumentNullException(nameof(dataSourceObjectList));
            _cancellationToken = cancellationToken;
            _progressReporter = progressReporter;

            var xpsDocumentSplicer = new XpsDocumentSplicer(reportProperties);
            Fill(ref xpsDocumentSplicer, dataSourceObjectList);
            if(IsCancellationRequested())
                return;

            SaveAsPdf(xpsDocumentSplicer, reportProperties.OutputDirectory);
            FinishProgress();
        }

        public static async Task GetPdfReportFromObjectAsync<T>(ReportProperties reportProperties, T dataSourceObject, CancellationToken token = default, IProgress<double> progress = null)
        {
            await GetPdfReportFromObjectListAsync(reportProperties, new List<T> { dataSourceObject }, token, progress).ConfigureAwait(false);
        }

        public static Task GetPdfReportFromObjectListAsync<T>(ReportProperties reportProperties, IEnumerable<T> dataSourceObjectList, CancellationToken token = default, IProgress<double> progress = null)
        {
            var taskCompletionSource = new TaskCompletionSource<object>();
            var thread = new Thread(() => GetPdfReportFromObjectList(reportProperties, dataSourceObjectList, token, progress));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return taskCompletionSource.Task;
        }

        public static FileStream GetPdfReportFromObjectList<T>(ReportProperties reportProperties, IEnumerable<T> dataSourceObjectList, CancellationToken cancellationToken = default, IProgress<double> progressReporter = null)
        {
            if(dataSourceObjectList == null)
                throw new ArgumentNullException(nameof(dataSourceObjectList));
            _cancellationToken = cancellationToken;
            _progressReporter = progressReporter;

            var xpsDocumentSplicer = new XpsDocumentSplicer(reportProperties);
            Fill(ref xpsDocumentSplicer, dataSourceObjectList);
            if(IsCancellationRequested())        
                return null;

            FileStream pdfFileStream = GetAsPdfFileStream(xpsDocumentSplicer);
            FinishProgress();
            return pdfFileStream;
        }

        private static void Fill<T>(ref XpsDocumentSplicer xpsDocumentSplicer, IEnumerable<T> dataSourceObjectList)
        {
            foreach(T sourceObject in dataSourceObjectList)
            {
                if(IsCancellationRequested())
                    break;
                xpsDocumentSplicer.AddXpsDocumentWith(sourceObject);
                ReportProgressFor(dataSourceObjectList);
            } 
        }

        private static bool IsCancellationRequested()
        {
            return _cancellationToken.IsCancellationRequested;
        }

        private static void ReportProgressFor<T>(IEnumerable<T> list)
        {
            double progress = CalculateProgressFor(list);
            _progressReporter?.Report(progress);
        }

        private static double CalculateProgressFor<T>(IEnumerable<T> list)
        {
            double listCount = list.Count();
            return _progress += (1 / listCount) * 85;  //Not *100 because after this processes is still some work left ;).
        }

        private static void SaveAsPdf(XpsDocumentSplicer xpsDocumentSplicer, string outputDirectory)
        {
            string tempXpsSavePath = GetTempXpsSavePath();
            xpsDocumentSplicer.SaveSplicedXpsDocumentTo(tempXpsSavePath);
            ConvertXpsToPdf(tempXpsSavePath, outputDirectory);
            File.Delete(tempXpsSavePath);
        }

        private static FileStream GetAsPdfFileStream(XpsDocumentSplicer xpsDocumentSplicer)
        {
            string tempXpsSavePath = GetTempXpsSavePath();
            string tempPdfSavePath = GetTempPdfSavePath();
            xpsDocumentSplicer.SaveSplicedXpsDocumentTo(tempXpsSavePath);
            ConvertXpsToPdf(tempXpsSavePath, tempPdfSavePath);
            FileStream pdfFileStream = File.Open(tempPdfSavePath, FileMode.OpenOrCreate);
            File.Delete(tempPdfSavePath);
            File.Delete(tempXpsSavePath);
            return pdfFileStream;
        }

        private static string GetTempXpsSavePath()
        {
            string verzeichnisName = AppDomain.CurrentDomain.BaseDirectory;
            return verzeichnisName + $"tempXpsFile{Directory.GetFiles(verzeichnisName).Length}.xps";
        }
        private static string GetTempPdfSavePath()
        {
            string verzeichnisName = AppDomain.CurrentDomain.BaseDirectory;
            return verzeichnisName + $"tempXpsFile{Directory.GetFiles(verzeichnisName).Length}.pdf";
        }

        private static void ConvertXpsToPdf(string xpsFileName, string outputDirectory)
        {
            using(var pdfXpsDoc = PdfSharp.Xps.XpsModel.XpsDocument.Open(xpsFileName))
            {
                PdfSharp.Xps.XpsConverter.Convert(pdfXpsDoc, outputDirectory, 0);
            }
        }

        private static void FinishProgress()
        {
            _progressReporter?.Report(100);
            _progress = 0;
            _cancellationToken = default; 
            _progressReporter = null;
            PimpedPaginator.GlobalPageCounter = 0;
        }

        #endregion
    }
}
