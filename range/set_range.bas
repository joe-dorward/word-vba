Sub set_range()
  ' sets the Range to be the 4th character in the document
  ' the selection merely shows that it has been set

  Set MyRange = ActiveDocument.Range(3, 4)
  MyRange.Select
  
End Sub
