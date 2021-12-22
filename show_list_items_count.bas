Sub show_list_items_count
  ' steps through the lists in the active-document
  ' MsgBox shows number of list-paragraphs in each list
  
  For Each List In ActiveDocument.Lists

    MsgBox List.ListParagraphs.Count

  Next List

End Sub
