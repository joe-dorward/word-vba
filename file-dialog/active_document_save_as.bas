Sub ActiveDocument_Save_As()

  Dim TheFileDialog As FileDialog
  Set TheFileDialog = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)

  Dim FullName As String ' includes path
  Dim Name As String
  
  Dim Path_To_Folder As String
  Dim Path_To_Folder_Length As Integer

  FullName = ActiveDocument.FullName
  Name = ActiveDocument.Name
  'MsgBox FullName & vbNewLine & Name
  
  Path_To_Folder_Length = Len(FullName) - Len(Name)
  Path_To_Folder = Left(FullName, Path_To_Folder_Length)
  'MsgBox Path_To_Folder
  
  TheFileDialog.InitialFileName = Path_To_Folder & Name
  TheFileDialog.Show
  TheFileDialog.Execute
  
End Sub
