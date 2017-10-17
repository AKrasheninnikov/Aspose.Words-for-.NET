// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using Aspose.Words.Markup;
using NUnit.Framework;
using System.IO;

namespace ApiExamples
{
    /// <summary>
    /// Tests that verify work with structured document tags in the document 
    /// </summary>
    [TestFixture]
    internal class ExStructuredDocumentTag : ApiExampleBase
    {
        [Test]
        public void RepeatingSection()
        {
            Document doc = new Document(MyDir + "TestRepeatingSection.docx");
            NodeCollection sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            //Assert that the node have sdttype - RepeatingSection and it's not detected as RichText
            StructuredDocumentTag sdt = (StructuredDocumentTag)sdts[0];
            Assert.AreEqual(SdtType.RepeatingSection, sdt.SdtType);

            //Assert that the node have sdttype - RichText 
            sdt = (StructuredDocumentTag)sdts[1];
            Assert.AreNotEqual(SdtType.RepeatingSection, sdt.SdtType);
        }

        [Test]
        public void CheckBox()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            sdtCheckBox.Checked = true;

            //Insert content control into the document
            builder.InsertNode(sdtCheckBox);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            NodeCollection sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            StructuredDocumentTag sdt = (StructuredDocumentTag)sdts[0];
            Assert.AreEqual(true, sdt.Checked);
        }

        [Test]
        public void CreatingCustomXml()
        {
            Document doc = new Document();

            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), "<root><text>Hello, World!</text></root>");

            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/text[1]", "");
            doc.FirstSection.Body.AppendChild(sdt);

            doc.Save(MyDir + @"\Artifacts\CustomXml Out.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\CustomXml Out.docx", MyDir + @"\Golds\CustomXml Gold.docx"));
        }

        [Test]
        public void ClearTextFromStructuredDocumentTags()
        {
            Document doc = new Document(MyDir + "TestRepeatingSection.docx");

            NodeCollection sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            Assert.IsNotNull(sdts);

            foreach (StructuredDocumentTag sdt in sdts)
            {
                sdt.Clear();
            }

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            Assert.AreEqual("Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.\r", sdts[0].GetText());
            Assert.AreEqual("Click here to enter text.\f", sdts[2].GetText());
        }
    }
}