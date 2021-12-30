
Sub Get_List_ListParagraph_ListType()
  ' get the list-type (of a) list-paragraph (of a) list
  ' for more on enumerated list-types, see: https://docs.microsoft.com/en-us/office/vba/api/word.wdlisttype
    
  Dim List_Number As Integer
  Dim ListParagraph_Number As Integer
    
  List_Number = 1
  ListParagraph_Number = 2
    
  ActiveDocument.Lists(List_Number).ListParagraphs.Item(ListParagraph_Number).Range.Select
  MsgBox ActiveDocument.Lists(List_Number).ListParagraphs.Item(ListParagraph_Number).Range.ListFormat.ListType
  Selection.Collapse
    
End Sub
