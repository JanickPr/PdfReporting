using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;

namespace PdfReporting.Logic
{
    /// <summary>
    /// The XpsDocumentManager is able to create ManagedXpsDocuments and manage their uris because 
    /// the XpsDocumentManager is used to combine a list of documents.
    /// </summary>
    public class XpsDocumentSplicer : List<ManagedXpsDocument>
    {
        private readonly ReportProperties _reportProperties;

        public XpsDocumentSplicer(ReportProperties reportProperties)
        {
            this._reportProperties = reportProperties;
        }

        public void AddXpsDocumentWith<T>(T dataSourceObject)
        {
            string bodyTemplateFilePath = GetBodyTemplateFilePathFrom(this._reportProperties.TemplateFolderPath);
            var managedFlowDocument = ManagedFlowDocument.GetFlowDocumentFrom(bodyTemplateFilePath, dataSourceObject);
            this.AddXpsDocumentWithContentFrom(managedFlowDocument);
        }

        private static string GetBodyTemplateFilePathFrom(string templateFolderPath)
        {
            try
            {
                List<string> folderContent = GetAllFileNamesFrom(templateFolderPath);
                return folderContent.First(fileName => fileName.EndsWith("BodyTemplate.xaml"));
            }
            catch(Exception)
            {
                throw new NullReferenceException("The folder does not contain any file ending with 'BodyTemplate.xaml'.");
            }
        }

        private void AddXpsDocumentWithContentFrom(ManagedFlowDocument managedFlowDocument)
        {
            ManagedXpsDocument managedXpsDocument = CreateManagedXpsDocumentFrom(managedFlowDocument);
            this.Add(managedXpsDocument);
        }

        private ManagedXpsDocument CreateManagedXpsDocumentFrom(ManagedFlowDocument managedFlowDocument)
        {
            XpsHeaderAndFooterDefinition xpsHeaderAndFooterDefinition = GetXpsHeaderAndFooterDefinitionWith(managedFlowDocument.DataContext);
            ManagedXpsDocument managedXpsDocument = GetNewManagedXpsDocument(xpsHeaderAndFooterDefinition);
            managedXpsDocument.CreateContentFromFlowDocument(managedFlowDocument);
            return managedXpsDocument;
        }

        private XpsHeaderAndFooterDefinition GetXpsHeaderAndFooterDefinitionWith(object dataContext)
        {
            string headerTemplateFilePath = GetHeaderTemplateFilePathFrom(this._reportProperties.TemplateFolderPath);
            string footerTemplateFilePath = GetFooterTemplateFilePathFrom(this._reportProperties.TemplateFolderPath);
            string bodyTemplateFilePath = GetBodyTemplateFilePathFrom(this._reportProperties.TemplateFolderPath);
            return new XpsHeaderAndFooterDefinition(headerTemplateFilePath, footerTemplateFilePath, dataContext);
        }

        private string GetHeaderTemplateFilePathFrom(string templateFolderPath)
        {
            try
            {
                List<string> folderContent = GetAllFileNamesFrom(templateFolderPath);
                return folderContent.First(fileName => fileName.EndsWith("HeaderTemplate.xaml"));
            }
            catch(InvalidOperationException)
            {
                return null;
            }
        }

        private string GetFooterTemplateFilePathFrom(string templateFolderPath)
        {
            try
            {
                List<string> folderContent = GetAllFileNamesFrom(templateFolderPath);
                return folderContent.First(fileName => fileName.EndsWith("FooterTemplate.xaml"));
            }
            catch(InvalidOperationException)
            {
                return null;
            }
        }

        private static List<string> GetAllFileNamesFrom(string templateFolderPath)
        {
            return Directory.GetFiles(templateFolderPath).ToList();
        }

        private ManagedXpsDocument GetNewManagedXpsDocument(XpsHeaderAndFooterDefinition xpsHeaderAndFooterDefinition)
        {
            Uri packageUri = GetNewIndexedPackageUri();
            Package package = GetNewPackageAt(packageUri);
            return new ManagedXpsDocument(packageUri, package, xpsHeaderAndFooterDefinition, this._reportProperties);
        }

        private Uri GetNewIndexedPackageUri()
        {
            return new Uri($"pack://temp{this.Count + 1}.xps");
        }

        private Package GetNewPackageAt(Uri packageUri)
        {
            if(packageUri == null)
                throw new ArgumentNullException(nameof(packageUri));
            return PackageHelper.OpenPackageInMemoryStream();
        }

        public void SaveSplicedXpsDocumentTo(string outputDirectory)
        {
            FixedDocumentSequence fixedDocumentSequence = GetSplicedFixedDocumentSequence();
            SaveXpsDocumentWithFixedDocumentSequenceTo(outputDirectory, fixedDocumentSequence);
            fixedDocumentSequence = null;
        }

        private FixedDocumentSequence GetSplicedFixedDocumentSequence()
        {
            List<PageContent> pages = this.GetAllPages();
            return CreateManagedXpsDocumentFromPages(pages);
        }

        private List<PageContent> GetAllPages()
        {
            var pages = new List<PageContent>();
            List<FixedDocument> fixedDocuments = GetAllFixedDocuments();
            fixedDocuments.ForEach(fixedDocument => pages.AddRange(fixedDocument.Pages));
            return pages;
        }

        private List<FixedDocument> GetAllFixedDocuments()
        {
            var allFixedDocuments = new List<FixedDocument>();

            foreach(ManagedXpsDocument xpsDocument in this)
            {
                IEnumerable<FixedDocument> fixedDocumentsOfXpsDocument = GetFixedDocumentsFrom(xpsDocument);
                allFixedDocuments.AddRange(fixedDocumentsOfXpsDocument);
            }

            return allFixedDocuments;
        }

        private IEnumerable<FixedDocument> GetFixedDocumentsFrom(ManagedXpsDocument xpsDocument)
        {
            FixedDocumentSequence fixedDocumentSequence = xpsDocument.GetFixedDocumentSequence();
            return fixedDocumentSequence.References.Select(r => r.GetDocument(true));
        }

        private FixedDocumentSequence CreateManagedXpsDocumentFromPages(IEnumerable<PageContent> pages)
        {
            var newSequence = new FixedDocumentSequence();
            var newDocReference = new DocumentReference();
            var newDoc = new FixedDocument();
            newDocReference.SetDocument(newDoc);

            foreach(PageContent page in pages)
            {
                var newPage = new PageContent
                {
                    Source = page.Source
                };
                (newPage as IUriContext).BaseUri = (page as IUriContext)?.BaseUri;
                newPage.GetPageRoot(true);
                newDoc.Pages.Add(newPage);
            }

            newSequence.References.Add(newDocReference);
            return newSequence;
        }

        private void SaveXpsDocumentWithFixedDocumentSequenceTo(string outputDirectory, FixedDocumentSequence fixedDocumentSequence)
        {
            File.Delete(outputDirectory);
            using(var xpsDocument = new XpsDocument(outputDirectory, FileAccess.ReadWrite))
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
