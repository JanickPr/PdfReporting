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
        private readonly ReportContentDefinition _xpsHeaderAndFooterDefinition;
        private readonly ReportProperties _reportProperties;

        public Uri PackageUri { get; set; }

        public ManagedXpsDocument(Uri packageUri, Package package, ReportContentDefinition xpsHeaderAndFooterDefinition, ReportProperties reportProperties)
            : base(package, CompressionOption.SuperFast, packageUri.ToString())
        {
            this.PackageUri = packageUri;
            RegisterAtPackageStoreWith(package, packageUri);
            this._xpsHeaderAndFooterDefinition = xpsHeaderAndFooterDefinition;
            this._reportProperties = reportProperties;
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
            PimpedPaginator documentPaginator = GetDocumentPaginatorFrom(managedFlowDocument);
            xpsSerializationManager.SaveAsXaml(documentPaginator);
        }

        private XpsSerializationManager GetXpsSerializationManager()
        {
            return new XpsSerializationManager(new XpsPackagingPolicy(this), false);
        }

        private PimpedPaginator GetDocumentPaginatorFrom(ManagedFlowDocument managedFlowDocument)
        {
            return new PimpedPaginator(managedFlowDocument, this._xpsHeaderAndFooterDefinition, this._reportProperties);
        }

        public new FixedDocumentSequence GetFixedDocumentSequence()
        {
            FixedDocumentSequence fixedDocumentSequence = base.GetFixedDocumentSequence();
            return fixedDocumentSequence ?? new FixedDocumentSequence();
        }

        private void UnregisterFromPackagestore(Uri packageUri)
        {
            PackageStore.RemovePackage(packageUri);
        }
    }
}
