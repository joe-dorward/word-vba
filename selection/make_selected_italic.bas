Sub Make_Selected_Italic()
  
  Selection.InsertBefore "<i>"
  Selection.InsertAfter "</i>"
  Selection.Font.Italic = False
  Selection.Collapse

End Sub
