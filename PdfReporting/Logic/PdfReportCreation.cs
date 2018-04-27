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
        private static int _progress;
        private static CancellationToken _cancellationToken;
        private static IProgress<int> _progressReporter;

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

        private static void CreatePdfReportFromObjectList<T>(ReportProperties reportProperties, IEnumerable<T> dataSourceObjectList, CancellationToken cancellationToken = default, IProgress<int> progressReporter = null)
        {
            if(dataSourceObjectList == null)
                throw new ArgumentNullException(nameof(dataSourceObjectList));
            _cancellationToken = cancellationToken;
            _progressReporter = progressReporter;

            var xpsDocumentSplicer = new XpsDocumentSplicer(reportProperties);
            Fill(ref xpsDocumentSplicer, dataSourceObjectList);
            CheckForCancellation();

                
            
            SaveAsPdf(xpsDocumentSplicer, reportProperties.OutputDirectory);
            FinishProgress();
        }

        private static void Fill<T>(ref XpsDocumentSplicer xpsDocumentSplicer, IEnumerable<T> dataSourceObjectList)
        {
            //ManagedThreadPool managedThreadPool = ManagedThreadPool.GetStaThreadPool();
            //managedThreadPool.FillWith(xpsDocumentSplicer.AddXpsDocumentWith, dataSourceObjectList);
            //managedThreadPool.Start();
            //ManagedThreadPool.WaitAll(new IWaitableResult[] { });
            //List<Thread> threadList = GetSTAThreadListFor(xpsDocumentSplicer.AddXpsDocumentWith, dataSourceObjectList);
            //Execute(threadList);
            //Await(threadList);
        }

        private static List<Thread> GetSTAThreadListFor<T>(Action<T> method, IEnumerable<T> parameterList)
        {
            var threadList = new List<Thread>();
            parameterList.ToList().ForEach(parameter =>
            {
                Thread thread = GetThreadFor(method, parameter);
                thread.SetApartmentState(ApartmentState.STA);
                threadList.Add(thread);
            });

            return threadList;
        }

        private static Thread GetThreadFor<T>(Action<T> method, T parameter)
        {
            return new Thread(() => method(parameter));
        }

        private static void Execute(List<Thread> threadList)
        {
            threadList.ForEach(thread => thread.Start());
        }

        private static void Await(List<Thread> threadList)
        {
            threadList.ForEach(thread =>
            {
                CheckForCancellation();
                thread.Join();
                ReportProgressFor(threadList);
            });
        }

        private static void CheckForCancellation()
        {
            if(_cancellationToken.IsCancellationRequested)
            {
                Thread.CurrentThread.Abort();
            }
        }

        private static void ReportProgressFor(IEnumerable<object> list)
        {
            int progress = CalculateProgressFor(list);
            _progressReporter?.Report(progress);
        }

        private static int CalculateProgressFor<T>(IEnumerable<T> list)
        {
            double listCount = list.Count();
            return _progress += (int)Math.Round((1 / listCount) * 90);  //*90 because after this processes is still some work left ;).
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

        private static void FinishProgress()
        {
            _progressReporter?.Report(100);
            _progressReporter = null;
            _cancellationToken = default;
            _progress = 0;
            PimpedPaginator.GlobalPageCounter = 0;
        }

        #endregion
    }
}
