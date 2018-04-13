using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.Windows.Xps.Serialization;

namespace PdfReporting.Logic
{
    public class ManagedXpsDocument : XpsDocument
    {
        private XpsHeaderAndFooterDefinition _xpsHeaderAndFooterDefinition;

        public Uri PackageUri { get; set; }

        public ManagedXpsDocument(Uri packageUri, Package package, XpsHeaderAndFooterDefinition xpsHeaderAndFooterDefinition) : base(package, CompressionOption.NotCompressed, packageUri.ToString())
        {
            PackageUri = packageUri;
            RegisterAtPackageStoreWith(package, packageUri);
            _xpsHeaderAndFooterDefinition = xpsHeaderAndFooterDefinition;
        }

        private void RegisterAtPackageStoreWith(Package package, Uri packageUri)
        {
            if (PackageStoreContains(packageUri))
            {
                UnregisterFromPackagestore(packageUri);
            }

            PackageStore.AddPackage(packageUri, package);
        }

        private bool PackageStoreContains(Uri packageUri)
        {
            return PackageStore.GetPackage(packageUri) != null;
        }

        public void CreateContentFromFlowDocument(ManagedFlowDocument managedFlowDocument)
        {
            XpsSerializationManager xpsSerializationManager = GetXpsSerializationManager();
            DocumentPaginator documentPaginator = GetDocumentPaginatorFrom(managedFlowDocument);
            xpsSerializationManager.SaveAsXaml(documentPaginator);
        }

        private XpsSerializationManager GetXpsSerializationManager()
        {
            return new XpsSerializationManager(new XpsPackagingPolicy(this), false);
        }

        private DocumentPaginator GetDocumentPaginatorFrom(ManagedFlowDocument managedFlowDocument)
        {
            DocumentPaginator paginator = ((IDocumentPaginatorSource)managedFlowDocument).DocumentPaginator;
            paginator = new DocumentPaginatorWrapper(paginator, new Size(796.8, 1123.2), default, _xpsHeaderAndFooterDefinition);
            return paginator;
        }

        public new FixedDocumentSequence GetFixedDocumentSequence()
        {
            FixedDocumentSequence fixedDocumentSequence = base.GetFixedDocumentSequence();
            return fixedDocumentSequence ?? new FixedDocumentSequence();
        }

        private List<DocumentReference> GetDocumentReferencesFrom(XpsDocument xpsDocument)
        {
            return xpsDocument.GetFixedDocumentSequence().References.ToList();
        }

        private XpsDocumentWriter GetXpsDocumentWriterFor(XpsDocument xpsDocument)
        {
            return XpsDocument.CreateXpsDocumentWriter(xpsDocument);
        }

        private void UnregisterFromPackagestore(Uri packageUri)
        {
            PackageStore.RemovePackage(packageUri);
        }
    }
}
