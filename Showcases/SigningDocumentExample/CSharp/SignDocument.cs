using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace SigningDocumentExample
{
    public class SignDocument : ApiExampleBase
    {
        public static void AddSignatureLineToDocument(DocumentBuilder builder, string signerPerson, string signerTitle, Guid signerId)
        {
            //Add some info about responsible person
            SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
            signatureLineOptions.Signer = signerPerson;
            signatureLineOptions.SignerTitle = signerTitle;

            //Add signature line for responsible person.
            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.Id = signerId;

            //Saves the document with signatureLine for for further sign
            builder.Document.Save(MyDir + "DocumentWithSignatureLine.doc");
        }
        
        /// <summary>
        /// Creates new document signed with responsible person.
        /// </summary>
        public static void SignDocumentWithPersonCertificate(Guid signerId, Byte[] signerImage, string pathToSignedDocument)
        {
            //Sign the document with responsible person certificate and image of its scanned signature.
            CertificateHolder certificateHolderSignerPerson = CreateCertificateHolder(MyDir + "morzal.pfx", "aw");

            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signerId;
            signOptions.SignatureLineImage = signerImage;

            DigitalSignatureUtil.Sign(MyDir + "DocumentWithSignatureLine.doc", pathToSignedDocument, certificateHolderSignerPerson, signOptions);
        }

        /// <summary>
        /// Creates certificate holder with responsible person certificate.
        /// </summary>
        private static CertificateHolder CreateCertificateHolder(string pathToCertificate, string certificatePassword)
        {
            return CertificateHolder.Create(pathToCertificate, certificatePassword);
        }
    }
}
