// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;

namespace SigningDocumentExample
{
    [TestFixture]
    public class SigningDocumentExample : ApiExampleBase
    {
        static readonly SignPersonList SignPersons = new SignPersonList();

        static readonly SignDocumentList SignDocuments = new SignDocumentList();

        static readonly string PathToSignedDocument = MyDir + "SignDocument.doc";

        /// <summary>
        /// Creates new document, signed by any persons from responsible divisions.
        /// </summary>
        [Test]
        public void FirstSigningDocument()
        {
            //Load scanned document.
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Get all required info about signer person
            string signerPosition = SignPersons.GetSignerPositionByName("Dhocs");
            Guid signerId = SignPersons.GetSignerIdByName("Dhocs");
            Byte[] signerImage = SignPersons.GetSignerImageByName("Dhocs");

            //Firts you need to add specific info about signer person
            SignDocument.AddSignatureLineToDocument(builder, "Dhocs", signerPosition, signerId);

            //Let it be signed by the 'Deputy Head of Corporate Services'.
            SignDocument.SignDocumentWithPersonCertificate(signerId, signerImage, PathToSignedDocument);

            //Get a signed document.
            Document signedDocument = new Document(PathToSignedDocument);

            //Write signed document into a database.
            SignDocuments.InsertDocument(signedDocument);
        }

        /// <summary>
        /// Signing document that are in database and update it.
        /// </summary>
        [Test]
        public void SigningDocumentFromDataTable()
        {
            //Load signed document from a data base.
            Byte[] signedByteArrayDocument = SignDocuments.GetDocumentByFileName(PathToSignedDocument);

            Document signedDocument = ConvertHepler.ConvertByteArrayToDocument(signedByteArrayDocument);
            DocumentBuilder builder = new DocumentBuilder(signedDocument);

            //Get all required info about signer person
            string signerPosition = SignPersons.GetSignerPositionByName("Dhocs");
            Guid signerId = SignPersons.GetSignerIdByName("Dhocs");
            Byte[] signerImage = SignPersons.GetSignerImageByName("Dhocs");

            //Firts you need to add specific info about signer person
            SignDocument.AddSignatureLineToDocument(builder, "Hocs", signerPosition, signerId);

            //Let it be signed by the 'Head of Corporate Services'.
            SignDocument.SignDocumentWithPersonCertificate(signerId, signerImage, PathToSignedDocument);

            //Get a signed document.
            signedDocument = new Document(PathToSignedDocument);

            //Write signed document into a database.
            SignDocuments.InsertDocument(signedDocument);
        }
    }
}