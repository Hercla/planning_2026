Attribute VB_Name = "Debug_Columns"
Sub ProbeColumns()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Feuille Personnel introuvable !"
        Exit Sub
    End If
    
    Dim msg As String
    msg = "Headers (Row 1):" & vbCrLf
    Dim i As Integer
    ' Check columns A to G
    For i = 1 To 7
        msg = msg & "Col " & i & " (" & Split(Cells(1, i).Address, "$")(1) & "): " & ws.Cells(1, i).value & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "Row 2 (First Data):" & vbCrLf
    For i = 1 To 7
        msg = msg & "Col " & i & ": " & ws.Cells(2, i).value & vbCrLf
    Next i
    
    MsgBox msg
End Sub
