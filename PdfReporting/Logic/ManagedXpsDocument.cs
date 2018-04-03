using System;
using System.Collections.Generic;
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
    public class ManagedXpsDocument : XpsDocument
    {
        public Uri PackageUri { get; set; }

        public ManagedXpsDocument(Uri packageUri, Package package) : base(package, CompressionOption.NotCompressed, packageUri.ToString())
        {
            PackageUri = packageUri;
            RegisterAtPackageStoreWith(package, packageUri);
        }

        private void RegisterAtPackageStoreWith(Package package, Uri packageUri)
        {
            if (PackageStoreContains(packageUri){
                UnregisterFromPackagestore(packageUri);
            }

            PackageStore.AddPackage(packageUri, package);
        }

        private bool PackageStoreContains(Uri packageUri)
        {
            return PackageStore.GetPackage(packageUri) != null;
        }

        public void CreateContentFromFlowDocument(FlowDocument flowDocument)
        {
            XpsSerializationManager xpsSerializationManager = GetXpsSerializationManager();
            DocumentPaginator documentPaginator = GetDocumentPaginatorFrom(flowDocument);
            xpsSerializationManager.SaveAsXaml(documentPaginator);
        }

        private XpsSerializationManager GetXpsSerializationManager()
        {
            return new XpsSerializationManager(new XpsPackagingPolicy(this), false);
        }

        private DocumentPaginator GetDocumentPaginatorFrom(FlowDocument flowDocument)
        {
            return ((IDocumentPaginatorSource)flowDocument).DocumentPaginator;
        }

        public void AddContentFrom(XpsDocument xpsDocument)
        {
            FixedDocumentSequence fixedDocumentSequence = GetCopyOfFixedDocumentSequenceFrom(xpsDocument);
            XpsDocumentWriter xpsDocumentWriter = GetXpsDocumentWriter();
            xpsDocumentWriter.Write(fixedDocumentSequence);            
        }

        public FixedDocumentSequence GetCopyOfFixedDocumentSequenceFrom(XpsDocument sourceXpsDocument)
        {
            FixedDocumentSequence fixedDocumentSequence = this.GetFixedDocumentSequence();
            List<DocumentReference> documentReferences = GetDocumentReferencesFrom(sourceXpsDocument);
            documentReferences.ForEach(documentReference => fixedDocumentSequence.AddCopyOf(documentReference));
            return fixedDocumentSequence;
        }

        private List<DocumentReference> GetDocumentReferencesFrom(XpsDocument xpsDocument)
        {
            return xpsDocument.GetFixedDocumentSequence().References.ToList();
        }

        private XpsDocumentWriter GetXpsDocumentWriter()
        {
            return XpsDocument.CreateXpsDocumentWriter(this);
        }

        private void UnregisterFromPackagestore(Uri packageUri)
        {
            PackageStore.RemovePackage(packageUri);
        }
    }
}
