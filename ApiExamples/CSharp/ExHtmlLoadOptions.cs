// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlLoadOptions : ApiExampleBase
    {
        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void SupportVml(bool supportVml)
        {
            //ExStart
            //ExFor:HtmlLoadOptions.SupportVml
            //ExSummary:Demonstrates how to parse html document with conditional comments like "<!--[if gte vml 1]>" and "<![if !vml]>"
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            //If SupportVml = true, then we parse "<!--[if gte vml 1]>", else parse "<![if !vml]>"
            loadOptions.SupportVml = supportVml;
            //Wait for a response, when loading external resources
            loadOptions.WebRequestTimeout = 1000;

            Document doc = new Document(MyDir + "Shape.VmlAndDml.htm", loadOptions);
            doc.Save(MyDir + @"\Artifacts\Shape.VmlAndDml.docx");
            //ExEnd
        }

        [Test]
        public void WebRequestTimeoutDefaultValue()
        {
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Assert.AreEqual(100000, loadOptions.WebRequestTimeout);
        }

        [Test]
        public void GetSelectAsSdt()
        {
            const string html = @"
                <html>
                    <select name='ComboBox' size='1'>
                        <option value='val1'>item1</option>
                        <option value='val2'></option>                        
                    </select>
                </html>
            ";

            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
            htmlLoadOptions.PreferredControlType = HtmlControlType.StructuredDocumentTag;
            
            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), htmlLoadOptions);
            NodeCollection nodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            
            StructuredDocumentTag tag = (StructuredDocumentTag)nodes[0];

            Assert.AreEqual(2, tag.ListItems.Count);

            Assert.AreEqual("val1", tag.ListItems[0].Value);
            Assert.AreEqual("val2", tag.ListItems[1].Value);
        }

        [Test]
        public void GetInputAsFormField()
        {
            const string html = @"
                <html>
                    <input type='text' value='Input value text' />
                </html>
            ";

            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
            htmlLoadOptions.PreferredControlType = HtmlControlType.FormField;

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), htmlLoadOptions);
            NodeCollection nodes = doc.GetChildNodes(NodeType.FormField, true);

            Assert.AreEqual(1, nodes.Count);

            FormField formField = (FormField)nodes[0];
            Assert.AreEqual("Input value text", formField.Result);
        }
    }
}