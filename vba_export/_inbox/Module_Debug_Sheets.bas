' Debug pour vérifier quelles feuilles sont trouvées
Sub Debug_Sheets()
    Dim wsCodesSpec As Worksheet
    Dim wsConfigCodes As Worksheet
    Dim msg As String
    
    On Error Resume Next
    Set wsCodesSpec = ThisWorkbook.Sheets("Codes_Speciaux")
    Set wsConfigCodes = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    
    msg = "=== VERIFICATION FEUILLES ===" & vbLf & vbLf
    
    If wsCodesSpec Is Nothing Then
        msg = msg & "Codes_Speciaux: NON TROUVE" & vbLf
    Else
        msg = msg & "Codes_Speciaux: OK (" & wsCodesSpec.Cells(wsCodesSpec.Rows.Count, "A").End(xlUp).Row & " lignes)" & vbLf
    End If
    
    If wsConfigCodes Is Nothing Then
        msg = msg & "Config_Codes: NON TROUVE" & vbLf
    Else
        msg = msg & "Config_Codes: OK (" & wsConfigCodes.Cells(wsConfigCodes.Rows.Count, "A").End(xlUp).Row & " lignes)" & vbLf
    End If
    
    msg = msg & vbLf & "=== TOUTES LES FEUILLES ===" & vbLf
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        msg = msg & "  - " & ws.Name & vbLf
    Next ws
    
    MsgBox msg, vbInformation, "Debug Feuilles"
End Sub
