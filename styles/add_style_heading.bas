Sub add_style_heading()
  ' adds the named-style "Heading" to the active-document
  ' on error - handles case of sub-procedure being run when "Heading" is already in the active-document

    On Error GoTo Add_Style
        If ActiveDocument.Styles("Heading").InUse Then
            Exit Sub
        End If

Add_Style:
    ActiveDocument.Styles.Add name:="Heading", Type:=wdStyleTypeParagraph
    With ActiveDocument.Styles("Heading").Font
        .Color = RGB(0, 175, 80) ' green
        .name = "Courier New"
        .Bold = False
        .Italic = False
    End With
End Sub
