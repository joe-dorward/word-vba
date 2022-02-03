Sub Make_Selected_Bold()  
  
  Selection.InsertBefore "<b>"
  Selection.InsertAfter "</b>"
  Selection.Font.Bold = False
  Selection.Collapse

End Sub
