' Item(1) is the first paragraph in the Paragraphs collection
Sub get_first_paragraph_text()

    MsgBox ActiveDocument.Paragraphs.Item(1).Range.Text

End Sub
