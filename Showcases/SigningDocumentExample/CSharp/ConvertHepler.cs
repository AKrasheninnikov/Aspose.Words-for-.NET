using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Words;

namespace SigningDocumentExample
{
    public class ConvertHepler : ApiExampleBase
    {
        /// <summary>
        /// Convert byte array to test image
        /// </summary>
        public static byte[] ConverTestImageToByteArray()
        {
            Image imageIn = Image.FromFile(MyDir + @"Images\LogoSmall.png");

            MemoryStream ms = new MemoryStream();
            imageIn.Save(ms, ImageFormat.Png);

            return ms.ToArray();
        }

        /// <summary>
        /// Convert byte array to AW document
        /// </summary>
        public static Document ConvertByteArrayToDocument(Byte[] document)
        {
            MemoryStream stream = new MemoryStream(document ?? throw new InvalidOperationException());
            Document documentFromDb = new Document(stream);

            return documentFromDb;
        }

        /// <summary>
        /// Convert AW document to byte array
        /// </summary>
        public static byte[] ConvertDocumentToByteArray(Document document)
        {
            MemoryStream documentArray = new MemoryStream();
            document.Save(documentArray, SaveFormat.Docx);

            return documentArray.ToArray();
        }
    }
}
