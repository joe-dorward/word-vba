Sub position_and_size()
  ' sets the left-edge of document-window to 10px from the left-edge of the screen
  ' sets the top-edge of document-window to 10px from the top-edge of the screen
  ' sets the width of the document-window to 800px
  ' sets the height of the document-window to 400px

  Application.Move Left:=10, Top:=10
  Application.Resize Width:=800, Height:=400

End Sub
