using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

namespace SigningDocumentExample
{
    public class SignDocumentList
    {
        public Guid Id { get; internal set; }
        public string FileName { get; internal set; }
        public byte[] Document { get; internal set; }

        public void InsertValue(SignDocumentList signDocument)
        {
            _signDocument.Add(signDocument);
        }

        public void InsertDocument(Document signedDocument)
        {
            byte[] byteArraySignedDocument = ConvertHepler.ConvertDocumentToByteArray(signedDocument);

            SignDocumentList dbSignedDocument = _signDocument.AsEnumerable().FirstOrDefault(p => p.FileName == signedDocument.OriginalFileName);

            //If this signed document are already exists, then we just update this document on new, else create new row in table.
            if (dbSignedDocument != null)
            {
                dbSignedDocument.Document = byteArraySignedDocument;
            }
            else
            {
                InsertValue(new SignDocumentList { Id = Guid.NewGuid(), FileName = signedDocument.OriginalFileName, Document = byteArraySignedDocument });
            }
        }

        public IEnumerable<SignDocumentList> GetAllSignDocuments()
        {
            return _signDocument;
        }

        public byte[] GetDocumentByFileName(string fileName)
        {
            return _signDocument.AsEnumerable().Where(p => p.FileName == fileName).Select(row => row.Document).FirstOrDefault();
        }

        private readonly List<SignDocumentList> _signDocument = new List<SignDocumentList>();
    }
}
