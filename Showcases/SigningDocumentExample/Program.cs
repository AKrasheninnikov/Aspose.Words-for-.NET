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
            // We need to create simple List with test signers for this example.
            CreateTestData();
            Console.WriteLine("Test data successfully added!");

            // Person who must sign the document
            string signPerson = "SignPerson 1";
            // Path to document that we need to signed
            string srcDocument = DataDir + "TestFile.docx";
            // Path to signed document
            string dstDocument = DataDir + "SignedDocument.docx";
            // Path to personal certificate
            string certifacate = DataDir + "morzal.pfx";
            // Password of the personal certificate
            string passwordCertificate = "aw";
            
            // Get signed document with personal certificate.
            Document signedDocument = GetSignedDocument(srcDocument, dstDocument, signPerson, certifacate, passwordCertificate);
            Console.WriteLine("Document successfully signed!");

            // Now we need add signed document to simple List.
            WriteSignedDocument(signedDocument, "New signed document");
            Console.WriteLine("Signed document successfully added!");
        }

        private static Document GetSignedDocument(string srcDocument, string dstDocument, string signPerson, string certificate, string passwordCertificate)
        {
            // Get document that we need to sign.
            Document baseDocument = new Document(srcDocument);
            DocumentBuilder builder = new DocumentBuilder(baseDocument);

            // Get sign person object by name of the person who must sign document.
            // This an example.
            // Actually, you need to return object from a data base.
            SignPerson signPersonInfo = (from c in mSignPersonList where c.Name == signPerson select c).FirstOrDefault();

            // Create holder of certificate instanse base on your personal certificate.
            // This is the test certificate generated for this example.
            CertificateHolder certificateHolder = CertificateHolder.Create(certificate, passwordCertificate);

            // Let's add signature to the document and sign it with personal certificate.
            SignDocument(builder, dstDocument, certificateHolder, signPersonInfo);

            return new Document(dstDocument);
        }

        /// <summary>
        /// Add signature line to the document and sign it with personal certificate
        /// </summary>
        /// <param name="builder">Class that provides methods for create SignatureLine</param>
        /// <param name="dstDocument">Path to signed document</param>
        /// <param name="certificateHolder">Holder of personal certificate instanse</param>
        /// <param name="signPersonInfo">SignPerson object which contains info about person who must sign document</param>
        private static void SignDocument(DocumentBuilder builder, string dstDocument, CertificateHolder certificateHolder, SignPerson signPersonInfo)
        {
            // Add info about responsible person who sign document.
            SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
            signatureLineOptions.Signer = signPersonInfo.Name;
            signatureLineOptions.SignerTitle = signPersonInfo.Position;

            // Add signature line for responsible person who sign document.
            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.Id = signPersonInfo.PersonId;

            // Save document with line signatures into temporary file for future signing.
            builder.Document.Save(dstDocument);

            // Link our signature line with personal signature.
            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signPersonInfo.PersonId;
            signOptions.SignatureLineImage = signPersonInfo.Image;

            // Sign document which contains signature line with personal certificate.
            DigitalSignatureUtil.Sign(dstDocument, dstDocument, certificateHolder, signOptions);
        }

        /// <summary>
        /// Example method for add signed document to simple List
        /// </summary>
        /// <param name="signedDocument">Signed document that we need to add</param>
        /// <param name="documentName">Name of the signed document</param>
        private static void WriteSignedDocument(Document signedDocument, string documentName)
        {
            // This just an example.
            // Actually, it will save or update object to data base.
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

        /// <summary>
        /// Create test data with info about sing persons
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