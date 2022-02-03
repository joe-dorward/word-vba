Sub Find_Bold()

  With Selection.Find
  
    .ClearFormatting
    .Font.Bold = True
    .Wrap = wdFindContinue
    .Execute
    
  End With
  
End Sub
