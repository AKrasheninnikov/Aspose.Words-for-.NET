// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExPclSaveOptions : ApiExampleBase
    {
        [Test]
        public void RasterizeElements()
        {
            Document doc = new Document(MyDir + "Document.EpubConversion.doc");

            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.RasterizeTransformedElements = true;

            doc.Save(MyDir + @"\Artifacts\Document.EpubConversion.pcl", saveOptions);
        }

        [Test]
        public void SetPrinterFont()
        {
            Document doc = new Document(MyDir + "Document.EpubConversion.doc");
            
            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.AddPrinterFont("Courier", "Courier");
            saveOptions.FalllbackFontName = "Times New Roman";

            doc.Save(MyDir + @"\Artifacts\Document.EpubConversion.pcl", saveOptions);
        }
    }
}