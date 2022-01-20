Sub blue_table_paragraphs()
  ' steps through the tables in the active-document
  ' sets text-color of table-paragraphs to blue
  
  For Each Table In ActiveDocument.Tables ' step through tables

    For Paragraph = 1 To (Table.Range.Paragraphs.Count - 1) ' step through paragraphs

      Table.Range.Paragraphs.Item(Paragraph).Range.Font.TextColor = RGB(0, 110, 190) ' Blue

    Next Paragraph

  Next Table

End Sub
