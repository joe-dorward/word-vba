' steps through each list in the document
Sub show_list_items_count
  
  For Each List In ActiveDocument.Lists

    MsgBox List.ListParagraphs.Count

  Next List

End Sub
