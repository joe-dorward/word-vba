Sub style_of_list_items
  ' steps through the lists in the active-document
  ' MsgBox shows the style-name of each list-paragraph, in each list
  
  For Each List In ActiveDocument.Lists
        
    For Index = 1 To (List.ListParagraphs.Count)
        
      MsgBox List.ListParagraphs.Item(Index).Range.Style
                   
    Next Index
        
  Next List
    
End Sub
