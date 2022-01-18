' this function returns the path to the active-document
Function Get_Path() As String

  Dim PathLength As Integer

  PathLength = Len(ActiveDocument.FullName) - Len(ActiveDocument.name)
  Get_Path = Left(ActiveDocument.FullName, PathLength)

End Function
