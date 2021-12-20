Sub ListType_Bullets
  ' Lists.Item(1) - is the first-list
  ' ListParagraphs.Item(1) - is first-paragraph
  ' MsgBox returns True if the first-paragraph in the first-list is bulleted

  MsgBox ActiveDocument.Lists.Item(1).ListParagraphs.Item(1).Range.ListFormat.ListType = wdListBullet
  
End Sub
