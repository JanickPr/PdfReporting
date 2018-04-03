using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Markup;

namespace PdfReporting.Logic
{
    public static class FixedDocumentSequenceExtensionMethods
    {
        public static void AddCopyOf(this FixedDocumentSequence fixedDocumentSequence, DocumentReference sourceDocumentReference)
        {
            DocumentReference documentReference = GetCopyOf(sourceDocumentReference);
            fixedDocumentSequence.References.Add(documentReference);
        }
        
        private static DocumentReference GetCopyOf(DocumentReference sourceDocumentReference)
        {
            DocumentReference documentReference = new DocumentReference();
            documentReference.Source = sourceDocumentReference.Source;
            (documentReference as IUriContext).BaseUri = (sourceDocumentReference as IUriContext).BaseUri;
            FixedDocument fixedDocument = sourceDocumentReference.GetDocument(true);
            documentReference.SetDocument(fixedDocument);
            return documentReference;
        }
    }
}
