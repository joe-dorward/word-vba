' This sub-procedure steps through the paragraphs of each table setting the TextColor of each
Sub Turn_Table_Paragraphs_Blue()

  For Each Table In ActiveDocument.Tables ' step through tables

    For Paragraph = 1 To (Table.Range.Paragraphs.Count - 1) ' step through paragraphs

      Table.Range.Paragraphs.Item(Paragraph).Range.Font.TextColor = RGB(0, 110, 190) ' Blue

    Next Paragraph

  Next Table

End Sub
