Sub Find_Italic()

  With Selection.Find
  
    .ClearFormatting
    .Font.Italic = True
    .Wrap = wdFindContinue
    .Execute
    
  End With
  
End Sub
