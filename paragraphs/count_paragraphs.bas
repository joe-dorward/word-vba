Sub count_paragraphs()
  ' steps through the Paragraphs collection, counting

    Dim ParagraphCounter As Integer
    ParagraphCounter = 0

    For Each Paragraph In ActiveDocument.Paragraphs
    
        ParagraphCounter = ParagraphCounter + 1
    
    Next Paragraph
    
    MsgBox ParagraphCounter
    
End Sub
