' Item(1) is the first paragraph in the Paragraphs collection
Sub set_first_paragraph_text()

    ActiveDocument.Paragraphs.Item(1).Range.Text = "This is the new text of the first paragraph." & vbNewLine

End Sub
