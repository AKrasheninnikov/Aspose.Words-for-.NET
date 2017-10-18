using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace SigningDocumentExample
{
    class SignDocumentsRepository : RepositoryBase
    {
        private static DataTable _dataTable;

        public SignDocumentsRepository(DataTable dataTable) : base(dataTable)
        {
            _dataTable = dataTable;
        }

        public new Object SearchFor(Func<DataRow, bool> predicate, Func<DataRow, object> selector)
        {
            return _dataTable.AsEnumerable().Where(predicate).Select(selector).FirstOrDefault();
        }

        public void InsertDocument(Document signedDocument)
        {
            //Create stream from signed document.
            MemoryStream stream = new MemoryStream();
            signedDocument.Save(stream, SaveFormat.Docx);

            byte[] byteArraySignedDocument = stream.ToArray();

            DataRow dbSignedDocument = SearchFor(p => p.Field<string>("FileName") == signedDocument.OriginalFileName).FirstOrDefault();

            //If this signed document are already exists, then we just update this document on new, else create new row in table.
            if (dbSignedDocument != null)
            {
                dbSignedDocument["Document"] = byteArraySignedDocument;
            }
            else
            {
                _dataTable.Rows.Add(Guid.NewGuid(), signedDocument.OriginalFileName, byteArraySignedDocument);
            }
        }

        public Byte[] GetSignDocument(string signedDocumentName)
        {
            return (byte[])this.SearchFor(p => p.Field<string>("FileName") == signedDocumentName, row => row.Field<Byte[]>("Document"));
        }
    }
}
