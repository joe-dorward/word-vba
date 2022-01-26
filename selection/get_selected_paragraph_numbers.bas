Sub Get_Selected_Paragraph_Numbers()

  Dim SelectionStart As Integer
  Dim SelectionEnd As Integer

  SelectionStart = Selection.Paragraphs.Item(1).Range.Information(wdFirstCharacterLineNumber)
  SelectionEnd = SelectionStart + Selection.Paragraphs.Count - 1 ' don't count the first-one twice
  
  MsgBox "You have selected paragraphs " & _
    "(" & SelectionStart & ") to " & _
    "(" & SelectionEnd & ")", , "Selection"

  ' do something
  
End Sub
