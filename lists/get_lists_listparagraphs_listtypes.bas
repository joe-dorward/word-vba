Sub get_lists_listparagraphs_listtypes()
  ' gets the list-type (of each) list-paragraph (of each) list
  ' for more on enumerated list-types, see: https://docs.microsoft.com/en-us/office/vba/api/word.wdlisttype
  
  For Each List In ActiveDocument.Lists
           
        For Each ListParagraph In List.ListParagraphs
        
            ListParagraph.Range.Select
            MsgBox ListParagraph.Range.ListFormat.ListType
            
        Next ListParagraph
            
    Next List
    Selection.Collapse
    
End Sub
