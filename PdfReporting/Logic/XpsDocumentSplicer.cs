using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.Windows.Xps.Serialization;

namespace PdfReporting.Logic
{
    /// <summary>
    /// The XpsDocumentManager is able to create ManagedXpsDocuments and manage their uris because 
    /// the XpsDocumentManager is used to combine a list of documents.
    /// </summary>
    public class XpsDocumentSplicer : List<ManagedXpsDocument>
    {
        public void AddXpsDocumentFrom<T>(string templateFilePath, T dataSourceObject)
        {
            FlowDocument flowDocument = new FlowDocument();
            flowDocument = flowDocument.InitializeFlowDocumentReportWith(templateFilePath, dataSourceObject);
            this.AddXpsDocumentWithContentFrom(flowDocument);
        }

        private void AddXpsDocumentWithContentFrom(FlowDocument flowDocument)
        {
            ManagedXpsDocument managedXpsDocument = CreateManagedXpsDocumentFromFlowDocument(flowDocument);
            this.Add(managedXpsDocument);
        }

        private ManagedXpsDocument CreateManagedXpsDocumentFromFlowDocument(FlowDocument flowDocument)
        {
            ManagedXpsDocument managedXpsDocument = GetNewManagedXpsDocument();
            managedXpsDocument.CreateContentFromFlowDocument(flowDocument);
            return managedXpsDocument;
        }

        private ManagedXpsDocument GetNewManagedXpsDocument()
        {
            Uri packageUri = GetNewIndexedPackageUri();
            Package package = GetNewPackageAt(packageUri);
            return new ManagedXpsDocument(packageUri, package);
        }

        private Uri GetNewIndexedPackageUri()
        {
            return new Uri($@"pack://temp{this.Count + 1}.xps");
        }

        private Package GetNewPackageAt(Uri packageUri)
        {
            Package package = PackageHelper.OpenPackageInMemoryStream();
            return package;
        }

        public void SaveSplicedXpsDocumentTo(String outputDirectory)
        {
            FixedDocumentSequence fixedDocumentSequence = GetSplicedFixedDocumentSequence();
            SaveXpsDocumentWithFixedDocumentSequenceTo(outputDirectory, fixedDocumentSequence);
        }

        private FixedDocumentSequence GetSplicedFixedDocumentSequence()
        {
            List<PageContent> pages = this.GetAllPages();
            FixedDocumentSequence fixedDocumentSequence = CreateManagedXpsDocumentFromPages(pages);
            return fixedDocumentSequence;
        }

        private List<PageContent> GetAllPages()
        {
            List<PageContent> pages = new List<PageContent>();
            List<FixedDocument> fixedDocuments = GetAllFixedDocuments();
            fixedDocuments.ForEach(fixedDocument => pages.AddRange(fixedDocument.Pages));
            return pages;
        }

        private List<FixedDocument> GetAllFixedDocuments()
        {
            List<FixedDocument> allFixedDocuments = new List<FixedDocument>();

            foreach (var xpsDocument in this)
            {
                IEnumerable<FixedDocument> fixedDocumentsOfXpsDocument = GetFixedDocumentsFrom(xpsDocument);
                allFixedDocuments.AddRange(fixedDocumentsOfXpsDocument);
            }

            return allFixedDocuments;
        }

        private IEnumerable<FixedDocument> GetFixedDocumentsFrom(ManagedXpsDocument xpsDocument)
        {
            FixedDocumentSequence fixedDocumentSequence = xpsDocument.GetFixedDocumentSequence();
            IEnumerable<FixedDocument> fixedDocuments = fixedDocumentSequence.References.Select(r => r.GetDocument(true));
            return fixedDocuments;
        }

        private FixedDocumentSequence CreateManagedXpsDocumentFromPages(IEnumerable<PageContent> pages)
        {
            FixedDocumentSequence newSequence = new FixedDocumentSequence();
            DocumentReference newDocReference = new DocumentReference();
            FixedDocument newDoc = new FixedDocument();
            newDocReference.SetDocument(newDoc);

            foreach (PageContent page in pages)
            {
                PageContent newPage = new PageContent();
                newPage.Source = page.Source;   
                (newPage as IUriContext).BaseUri = ((IUriContext)page).BaseUri;
                newPage.GetPageRoot(true);
                newDoc.Pages.Add(newPage);
            }

            // The order in which you do this is REALLY important
            newSequence.References.Add(newDocReference);

            return newSequence;
        }

        private void SaveXpsDocumentWithFixedDocumentSequenceTo(string outputDirectory, FixedDocumentSequence fixedDocumentSequence)
        {
            File.Delete(outputDirectory);
            using (XpsDocument xpsDocument = new XpsDocument(outputDirectory, FileAccess.ReadWrite))
            {
                XpsDocumentWriter xpsDocumentWriter = GetXpsDocumentWriterFor(xpsDocument);
                xpsDocumentWriter.Write(fixedDocumentSequence);
            }
        }

        private XpsDocumentWriter GetXpsDocumentWriterFor(XpsDocument xpsDocument)
        {
            return XpsDocument.CreateXpsDocumentWriter(xpsDocument);
        }
    }
}
