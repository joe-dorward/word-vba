Sub example_select_case()
  ' example of select-case

  Dim Number As Integer
    
  Number = InputBox("Enter a positive-number less-than 5", "Select Case")
    
  Select Case Number
    
      Case 1
        MsgBox "One"
    
      Case 2
        MsgBox "Two"
 
      Case 3
        MsgBox "Three"
    
      Case 4
        MsgBox "Four"

      Case Else
        MsgBox "Out of range"
        
  End Select

End Sub
