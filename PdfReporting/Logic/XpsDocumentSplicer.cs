using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
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
        public void AddFromFlowDocument(FlowDocument flowDocument)
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

        public ManagedXpsDocument Splice()
        {
            ManagedXpsDocument xpsDocument = GetNewManagedXpsDocument();
            this.ForEach(sourceXpsDocument => xpsDocument.AddContentFrom(sourceXpsDocument));
            
            return xpsDocument;
        }
    }
}
