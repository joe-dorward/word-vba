' GLOBAL VARIABLES
Dim Start_Time As Date
Dim End_Time As Date
Dim Processing_Time As Date
' ---------- ---------- ---------- ---------- ----------
Sub how_long()
    ' reports the difference between Start_Time and End_Time

    ' get Start_Time
    Start_Time = Time

    ' some process
    For T = 1 To 1000000000
    
    Next T
    
    ' get End_Time
    End_Time = Time

    ' calculate Processing_Time
    Processing_Time = End_Time - Start_Time

    ' report
    MsgBox "Start Time: " & vbTab & Start_Time & vbLf & _
    "End Time: " & vbTab & End_Time & vbLf & _
    "Processing Time: " & vbTab & Int(CSng(Processing_Time * 24 * 60)) & ":" & Format(Processing_Time, "ss"), , "how_long()"

End Sub
