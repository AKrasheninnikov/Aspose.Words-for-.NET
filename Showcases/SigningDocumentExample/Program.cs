// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace SigningDocumentExample
{
    [TestFixture]
    public class Program
    {
        // Sample infrastructure.
        static readonly string ExeDir = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath) + Path.DirectorySeparatorChar;
        static readonly string DataDir = new Uri(new Uri(ExeDir), @"../../Data/").LocalPath;
        static readonly string TestImage = DataDir + @"Images\LogoSmall.png";

        [Test]
        public static void Main()
        {
            // We need to create simple list with test signers for this example.
            CreateTestData();
            Console.WriteLine("Test data successfully added!");

            // Person who must sign the document
            string signPerson = "SignPerson 1";
            // Path to document that we need to sign
            string srcDocument = DataDir + "TestFile.docx";
            // Path to the signed document
            string dstDocument = DataDir + "SignedDocument.docx";
            // Path to the personal certificate
            string certifacate = DataDir + "morzal.pfx";
            // Password of the personal certificate
            string passwordCertificate = "aw";
            
            // Get signed document with a personal certificate.
            Document signedDocument = GetSignedDocument(srcDocument, dstDocument, signPerson, certifacate, passwordCertificate);
            Console.WriteLine("Document successfully signed!");

            // Now we need add signed document to simple List.
            WriteSignedDocument(signedDocument, "New signed document");
            Console.WriteLine("Signed document successfully added!");
        }

        /// <summary>
        /// Get signed document 
        /// </summary>
        /// <param name="srcDocument">Path to the document to sign</param>
        /// <param name="dstDocument">Path to the signed document</param>
        /// <param name="signPerson">Person name who need to sign document</param>
        /// <param name="certificate">Path to the personal certificate</param>
        /// <param name="passwordCertificate">Password of the personal certificate</param>
        /// <returns>Returns signed document</returns>
        private static Document GetSignedDocument(string srcDocument, string dstDocument, string signPerson, string certificate, string passwordCertificate)
        {
            // Get document that we need to sign.
            Document baseDocument = new Document(srcDocument);
            DocumentBuilder builder = new DocumentBuilder(baseDocument);

            // Get sign person object by name of the person who must sign a document.
            // This an example.
            // Actually, you need to return object from a data base.
            SignPerson signPersonInfo = (from c in mSignPersonList where c.Name == signPerson select c).FirstOrDefault();

            // Create holder of certificate instance base on your personal certificate.
            // This is the test certificate generated for this example.
            CertificateHolder certificateHolder = CertificateHolder.Create(certificate, passwordCertificate);

            // Let's add signature to the document and sign it with a personal certificate.
            SignDocument(builder, dstDocument, certificateHolder, signPersonInfo);

            return new Document(dstDocument);
        }

        /// <summary>
        /// Add signature line to the document and sign it with personal certificate
        /// </summary>
        /// <param name="builder">Class that provides methods for create SignatureLine</param>
        /// <param name="dstDocument">Path to the signed document</param>
        /// <param name="certificateHolder">Holder of personal certificate instance</param>
        /// <param name="signPersonInfo">SignPerson object which contains info about person who must sign a document</param>
        private static void SignDocument(DocumentBuilder builder, string dstDocument, CertificateHolder certificateHolder, SignPerson signPersonInfo)
        {
            // Add info about responsible person who sign a document.
            SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
            signatureLineOptions.Signer = signPersonInfo.Name;
            signatureLineOptions.SignerTitle = signPersonInfo.Position;

            // Add signature line for responsible person who sign a document.
            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.Id = signPersonInfo.PersonId;

            // Save a document with line signatures into temporary file for future signing.
            builder.Document.Save(dstDocument);

            // Link our signature line with personal signature.
            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signPersonInfo.PersonId;
            signOptions.SignatureLineImage = signPersonInfo.Image;

            // Sign a document which contains signature line with personal certificate.
            DigitalSignatureUtil.Sign(dstDocument, dstDocument, certificateHolder, signOptions);
        }

        /// <summary>
        /// Example method for adding or updating signed document into simple list
        /// </summary>
        /// <param name="signedDocument">Signed document</param>
        /// <param name="documentName">Name of the signed document</param>
        private static void WriteSignedDocument(Document signedDocument, string documentName)
        {
            // This just an example.
            // Actually, it will save or update object to data base.
            SignDocument existingDocument = (from c in mSignDocumentList where c.DocumentName == documentName select c).FirstOrDefault();

            if (existingDocument != null)
            {
                existingDocument.Document = ConvertHepler.ConvertDocumentToByteArray(signedDocument);
            }
            else
            {
                mSignDocumentList = new List<SignDocument>
                {
                    new SignDocument
                    {
                        DocumentId = Guid.NewGuid(),
                        DocumentName = documentName,
                        Document = ConvertHepler.ConvertDocumentToByteArray(signedDocument)
                    }
                };
            }
        }

        /// <summary>
        /// Create test data contains info about sing persons
        /// </summary>
        private static void CreateTestData()
        {
            mSignPersonList = new List<SignPerson>
            {
                new SignPerson
                {
                    PersonId = Guid.NewGuid(),
                    Name = "SignPerson 1",
                    Position = "Head of Department",
                    Image = ConvertHepler.ConverImageToByteArray(TestImage)
                },
                new SignPerson
                {
                    PersonId = Guid.NewGuid(),
                    Name = "SignPerson 2",
                    Position = "Deputy Head of Department",
                    Image = ConvertHepler.ConverImageToByteArray(TestImage)
                }
            };
        }

        private static List<SignPerson> mSignPersonList;
        private static List<SignDocument> mSignDocumentList;
    }
}