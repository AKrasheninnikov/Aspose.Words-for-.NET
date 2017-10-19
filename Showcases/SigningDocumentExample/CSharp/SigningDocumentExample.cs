// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace SigningDocumentExample
{
    [TestFixture]
    public class SigningDocumentExample : ApiExampleBase
    {
        static readonly DataTable SignPersonsTable = CreateSignPersonsTable();

        static readonly DataTable SignDocumentsTable = CreateSignDocumentsTable();

        //Get access to sign persons repository methods
        static readonly SignPersonsRepository SignPersonsRepo = new SignPersonsRepository(SignPersonsTable);

        //Get access to sign documents repository methods
        static readonly SignDocumentsRepository SignDocumentsRepo = new SignDocumentsRepository(SignDocumentsTable);

        /// <summary>
        /// Creates new document, signed by any persons from responsible divisions.
        /// </summary>
        [Test]
        public void FirstSigningDocument()
        {
            //Load scanned document.
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Set path to the document that will be signed.
            string pathToSignedDocument = MyDir + "SignDocument.doc";

            //Get all required info about signer person
            string signerPosition = SignPersonsRepo.GetSignerPositionByName("Dhocs");
            Guid signerId = SignPersonsRepo.GetSignerIdByName("Dhocs");
            Byte[] signerImage = SignPersonsRepo.GetSignerImageByName("Dhocs");
            
            //Firts you need to add specific info about signer person
            SignDocument.AddSignatureLineToDocument(builder, "Dhocs", signerPosition, signerId);

            //Let it be signed by the 'Deputy Head of Corporate Services'.
            SignDocument.SignDocumentWithPersonCertificate(signerId, signerImage, pathToSignedDocument);

            //Get a signed document.
            Document signedDocument = new Document(pathToSignedDocument);

            //Write signed document into a database.
            SignDocumentsRepo.InsertDocument(signedDocument);
        }

        /// <summary>
        /// Signing document that are in database and update it.
        /// </summary>
        [Test]
        public void SigningDocumentFromDataTable()
        {
            string pathToSignedDocument = MyDir + "SignDocument.doc";

            //Load signed document from a data base.
            Byte[] signedByteArrayDocument = SignDocumentsRepo.GetSignDocument(pathToSignedDocument);

            Document signedDocument = ConvertByteArrayToDocument(signedByteArrayDocument);
            DocumentBuilder builder = new DocumentBuilder(signedDocument);

            //Get all required info about signer person
            string signerPosition = SignPersonsRepo.GetSignerPositionByName("Dhocs");
            Guid signerId = SignPersonsRepo.GetSignerIdByName("Dhocs");
            Byte[] signerImage = SignPersonsRepo.GetSignerImageByName("Dhocs");

            //Firts you need to add specific info about signer person
            SignDocument.AddSignatureLineToDocument(builder, "Hocs", signerPosition, signerId);

            //Let it be signed by the 'Head of Corporate Services'.
            SignDocument.SignDocumentWithPersonCertificate(signerId, signerImage, pathToSignedDocument);

            //Get a signed document.
            signedDocument = new Document(pathToSignedDocument);

            //Write signed document into a database.
            SignDocumentsRepo.InsertDocument(signedDocument);
        }

        /// <summary>
        /// Get image bytes for saving into a database.
        /// </summary>
        private static byte[] ConverImageToByteArray(Image imageIn)
        {
            MemoryStream ms = new MemoryStream();
            imageIn.Save(ms, ImageFormat.Png);

            return ms.ToArray();
        }

        /// <summary>
        /// Convert byte array to AW document
        /// </summary>
        private static Document ConvertByteArrayToDocument(Byte[] document)
        {
            MemoryStream stream = new MemoryStream(document ?? throw new InvalidOperationException());
            Document documentFromDb = new Document(stream);

            return documentFromDb;
        }

        /// <summary>
        /// Creates new table with responsible persons.
        /// </summary>
        private static DataTable CreateSignPersonsTable()
        {
            DataTable table = new DataTable("SignPersons");
                        
            DataColumn column = new DataColumn("Id", typeof(Guid));
            column.Unique = true;
            table.Columns.Add(column);

            column = new DataColumn("Name", typeof(String));
            table.Columns.Add(column);

            column = new DataColumn("Position", typeof(String));
            table.Columns.Add(column);

            column = new DataColumn("Image", typeof(Byte[]));
            table.Columns.Add(column);

            Image img = Image.FromFile(MyDir + @"Images\LogoSmall.png");

            table.Rows.Add(Guid.Parse("CDAA3044-8017-4E07-BFF4-93EA14A3A6C9"), "Hocs", "Head of Corporate Services", ConverImageToByteArray(img));
            table.Rows.Add(Guid.Parse("1C22DFF1-B98E-4F65-888F-D55F9A968CD3"), "Dhocs", "Deputy Head of Corporate Services", ConverImageToByteArray(img));
            table.Rows.Add(Guid.Parse("C1DE3C2C-B2F8-4952-96BE-D300FBB9D26B"), "Hofd", "Head of Finance department", ConverImageToByteArray(img));
            table.Rows.Add(Guid.Parse("0A31FD51-46AF-4600-A188-64887881EC47"), "Hoad", "Head of Accounts department", ConverImageToByteArray(img));
            table.Rows.Add(Guid.Parse("409DB46E-4172-4679-ACDE-D78DDB18214F"), "Hdoit", "Head Department of IT", ConverImageToByteArray(img));

            return table;
        }

        /// <summary>
        /// Creates table for saving signed documents.
        /// </summary>
        private static DataTable CreateSignDocumentsTable()
        {
            DataTable table = new DataTable("SignDocuments");

            DataColumn column = new DataColumn("Id", typeof(Guid));
            column.Unique = true;
            table.Columns.Add(column);

            column = new DataColumn("FileName", typeof(string));
            table.Columns.Add(column);

            column = new DataColumn("Document", typeof(Byte[]));
            table.Columns.Add(column);

            return table;
        }
    }
}