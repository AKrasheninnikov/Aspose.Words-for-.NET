'
' Simple test with image shape
'
Sub TestWrapper5
    Set WDocumentFactory = CreateObject("Aspose.Words.Wrapper.WDocumentFactory")
    Set WObjectFactory = CreateObject("Aspose.Words.Wrapper.WObjectFactory")

    Set WDoc = WDocumentFactory.OpenFromFile("TestWrapper.docx")

    Set WBuilder = CreateObject("Aspose.Words.Wrapper.WDocumentBuilder")
    Set WBuilder.Document = WDoc

    Set Node = CreateObject("Aspose.Words.Wrapper.WNodeType")
    Set WImage = WDoc.GetChildNode(Node.GetShapeNodeType())
    
    MsgBox("IsImage=" & WImage.IsImage)
End Sub

Call TestWrapper5