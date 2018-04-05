using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
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

        public static void SaveAsXps<T>(string templateFileName, IEnumerable<T> dataSourceList, string outputDirectory)
        {
            XpsDocumentSplicer xpsDocumentSplicer = new XpsDocumentSplicer();

            //foreach (var dataSourceItem in dataSourceList)
            //{
            FlowDocument flowDocument = GetFlowDocumentFrom(templateFileName);
            flowDocument.SetUpDataContext(dataSourceList.ElementAt(0));
            xpsDocumentSplicer.AddXpsDocumentWithContentFrom(flowDocument);
            //}

            xpsDocumentSplicer.SaveSplicedXpsDocumentTo(outputDirectory);
        }

        public static void SaveAsXps<T>(string templateFileName, T dataSource, string outputDirectory)
        {
            FlowDocument flowDocument = GetFlowDocumentFrom(templateFileName);
            flowDocument.SetUpDataContext(dataSource);
            XpsDocumentSplicer xpsDocumentSplicer = new XpsDocumentSplicer();
            xpsDocumentSplicer.AddXpsDocumentWithContentFrom(flowDocument);
            xpsDocumentSplicer.SaveSplicedXpsDocumentTo(outputDirectory);
        }

        private static FlowDocument GetFlowDocumentFrom(string templateFilePath)
        {
            FlowDocument flowDocument = new FlowDocument();
            flowDocument = flowDocument.LoadFromTemplate(templateFilePath);
            return flowDocument;
        }

        

        public static void MergeXpsDocument(string outputDirectory, List<XpsDocument> sourceDocuments)
        {
            XpsDocument xpsDocument = new XpsDocument(outputDirectory, System.IO.FileAccess.ReadWrite);
            XpsDocumentWriter xpsDocumentWriter = XpsDocument.CreateXpsDocumentWriter(xpsDocument);
            FixedDocumentSequence fixedDocumentSequence = new FixedDocumentSequence();

            foreach (XpsDocument doc in sourceDocuments)
            {
                FixedDocumentSequence sourceSequence = doc.GetFixedDocumentSequence();
                foreach (DocumentReference dr in sourceSequence.References)
                {
                    DocumentReference newDocumentReference = new DocumentReference();
                    newDocumentReference.Source = dr.Source;
                    (newDocumentReference as IUriContext).BaseUri = (dr as IUriContext).BaseUri;
                    FixedDocument fd = newDocumentReference.GetDocument(true);
                    newDocumentReference.SetDocument(fd);
                    fixedDocumentSequence.References.Add(newDocumentReference);
                }
            }
            xpsDocumentWriter.Write(fixedDocumentSequence);
            xpsDocument.Close();
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
            pdfDokument.Save(outputDirectory);


            pdfDokument.Dispose();
            tempPdfDocument.Dispose();


            //Hier wird die leere erste Seite wieder entfernt, die in ErstellePdfInhalt<T> erstellt wurde.
            //var pdfDokument = new PdfSharp.Pdf.PdfDocument(outputDirectory);
            //pdfDokument.Pages.RemoveAt(0);
            //pdfDokument.Save(outputDirectory);
        }
        #endregion
    }
}
