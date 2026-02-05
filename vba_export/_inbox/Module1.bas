' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module1"
Public Sub AllerFeuilConfig_Safe()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Feuil_Config")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La feuille Feuil_Config n'existe pas.", vbExclamation
    Else
        ws.Activate
    End If
End Sub
Public Sub OpenConfig()
    ThisWorkbook.Worksheets("Feuil_Config").Activate
End Sub

