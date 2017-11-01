// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDigitalSignatureUtil : ApiExampleBase
    {
        [Test]
        public void RemoveAllSignatures()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
            //ExSummary:Shows how to remove every signature from a document.
            //By stream:
            Stream docStreamIn = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            Stream docStreamOut = new FileStream(MyDir + @"\Artifacts\Document.NoSignatures.FromStream.docx", FileMode.Create);

            DigitalSignatureUtil.RemoveAllSignatures(docStreamIn, docStreamOut);

            Document docWithoutSignatures = new Document(docStreamOut); //ExSkip

            DigitalSignatureCollection signatures = docWithoutSignatures.DigitalSignatures; //ExSkip

            if (signatures.Count > 0) //ExSkip
            { //ExSkip
                Assert.Fail("The document have signatures"); //ExSkip
            } //ExSkip

            docStreamIn.Close();
            docStreamOut.Close();

            //By string:
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outFileName = MyDir + @"\Artifacts\Document.NoSignatures.FromString.docx";

            DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outFileName);
            //ExEnd

            docWithoutSignatures = new Document(outFileName); //ExSkip

            signatures = docWithoutSignatures.DigitalSignatures; //ExSkip

            if (signatures.Count > 0) //ExSkip
            { //ExSkip
                Assert.Fail("The document have signatures"); //ExSkip
            } //ExSkip
        }

        [Test]
        public void LoadSignatures()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
            //ExFor:DigitalSignatureUtil.LoadSignatures(String)
            //ExSummary:Shows how to load signatures from a document by stream and by string.
            Stream docStream = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);

            // By stream:
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.LoadSignatures(docStream);
            docStream.Close();

            // By string:
            digitalSignatures = DigitalSignatureUtil.LoadSignatures(MyDir + "Document.DigitalSignature.docx");
            //ExEnd
        }

        [Test]
        public void SignDocument()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
            //ExFor:SignOptions.Comments
            //ExFor:SignOptions.SignTime
            //ExSummary:Shows how to sign documents.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //By String:
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outputDocFileName = MyDir + @"\Artifacts\Document.DigitalSignature.docx";

            SignOptions signOptions = new SignOptions();
            signOptions.Comments = "My comment";
            signOptions.SignTime = DateTime.Now;

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, certificateHolder, signOptions);

            //By Stream:
            Stream docInStream = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            Stream docOutStream = new FileStream(MyDir + @"\Artifacts\Document.DigitalSignature.docx", FileMode.OpenOrCreate);

            DigitalSignatureUtil.Sign(docInStream, docOutStream, certificateHolder, signOptions);
            //ExEnd

            docInStream.Dispose();
            docOutStream.Dispose();
        }

        [Test]
        public void IncorrectPasswordForDecrypring()
        {
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");
            
            Document doc = new Document(MyDir + "Document.Encrypted.docx", new LoadOptions("docPassword"));
            string outputDocFileName = MyDir + @"\Artifacts\Document.Encrypted.docx";

            SignOptions signOptions = new SignOptions();
            signOptions.Comments = "Comment";
            signOptions.SignTime = DateTime.Now;
            signOptions.DecryptionPassword = "docPassword1";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            Assert.That(() => DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, certificateHolder, signOptions), Throws.TypeOf<IncorrectPasswordException>(), "The document password is incorrect.");
        }

        [Test]
        public void SingDocumentWithPasswordDecrypring()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
            //ExFor:SignOptions.DecryptionPassword
            //ExSummary:Shows how to sign encrypted documents
            // Create certificate holder from a file.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //By String:
            Document doc = new Document(MyDir + "Document.Encrypted.docx", new LoadOptions("docPassword"));
            string outputDocFileName = MyDir + @"\Artifacts\Document.Encrypted.docx";

            SignOptions signOptions = new SignOptions();
            signOptions.Comments = "Comment";
            signOptions.SignTime = DateTime.Now;
            signOptions.DecryptionPassword = "docPassword";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, certificateHolder, signOptions);

            // Open encrypted document from a file.
            Document signedDoc = new Document(outputDocFileName, new LoadOptions("docPassword"));

            // Check that encrypted document was successfully signed.
            DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
            if (signatures.IsValid && (signatures.Count > 0))
            {
                Assert.Pass(); //The document was signed successfully
            }
        }

        [Test] //ExSkip
        public void SingStreamDocumentWithPasswordDecrypring() //ExSkip
        {
            // Create certificate holder from a file.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //By Stream:
            Stream docInStream = new FileStream(MyDir + "Document.Encrypted.docx", FileMode.Open);
            Stream docOutStream = new FileStream(MyDir + @"\Artifacts\Document.Encrypted.docx", FileMode.OpenOrCreate);

            SignOptions signOptions = new SignOptions();
            signOptions.Comments = "Comment";
            signOptions.SignTime = DateTime.Now;
            signOptions.DecryptionPassword = "docPassword";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(docInStream, docOutStream, certificateHolder, signOptions);

            // Open encrypted document from a file.
            Document signedDoc = new Document(docOutStream, new LoadOptions("docPassword"));

            // Check that encrypted document was successfully signed.
            DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
            if (signatures.IsValid && (signatures.Count > 0))
            {
                docInStream.Dispose();
                docOutStream.Dispose();

                Assert.Pass(); //The document was signed successfully
            }
            //ExEnd
        }

        [Test]
        public void NoArgumentsForSing()
        {
            SignOptions signOptions = new SignOptions();
            signOptions.Comments = String.Empty;
            signOptions.SignTime = DateTime.Now;
            signOptions.DecryptionPassword = String.Empty;

            Assert.That(() => DigitalSignatureUtil.Sign(String.Empty, String.Empty, null, signOptions), Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void NoCertificateForSign()
        {
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outputDocFileName = MyDir + @"\Artifacts\Document.DigitalSignature.docx";

            SignOptions signOptions = new SignOptions();
            signOptions.Comments = "Comment";
            signOptions.SignTime = DateTime.Now;
            signOptions.DecryptionPassword = "docPassword";

            Assert.That(() => DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, null, signOptions), Throws.TypeOf<ArgumentNullException>());
        }
    }
}