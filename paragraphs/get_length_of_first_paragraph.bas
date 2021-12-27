' len() returns the string-length
Sub get_length_of_first_paragraph()

    MsgBox Len(ActiveDocument.Paragraphs.Item(1).Range.Text) ' includes paragraph-mark
    
    MsgBox Len(ActiveDocument.Paragraphs.Item(1).Range.Text) - 1 ' excludes paragraph-mark

End Sub
