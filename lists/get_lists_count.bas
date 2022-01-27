Sub Get_Lists_Count()

  Dim ListsCount As Integer
  
  ListsCount = ActiveDocument.Lists.Count
  
  MsgBox ListsCount, , "Count"
  
  ' do something
  
End Sub
