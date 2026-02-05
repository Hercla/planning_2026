Attribute VB_Name = "Debug_DictCodes"
'Attribute VB_Name = "Debug_DictCodes"
Option Explicit

Sub Debug_Voir_DictCodes()
    Dim wsConfigCodes As Worksheet
    On Error Resume Next
    Set wsConfigCodes = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    
    If wsConfigCodes Is Nothing Then
        MsgBox "Config_Codes introuvable!", vbCritical
        Exit Sub
    End If
    
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    Dim lr As Long: lr = wsConfigCodes.Cells(wsConfigCodes.Rows.count, "A").End(xlUp).row
    Dim arr As Variant: arr = wsConfigCodes.Range("A2:O" & lr).value
    
    Dim rapport As String
    rapport = "=== CODES C15/C19/C20 DANS CONFIG_CODES ===" & vbCrLf & vbCrLf
    
    Dim i As Long, code As String
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        Dim codeUp As String: codeUp = UCase(Replace(code, " ", ""))
        
        If codeUp Like "*C15*" Or codeUp Like "*C19*" Or codeUp Like "*C20*" Then
            rapport = rapport & "Ligne " & (i + 1) & ": Code=[" & code & "]"
            rapport = rapport & " Matin=" & arr(i, 12)
            rapport = rapport & " PM=" & arr(i, 13)
            rapport = rapport & " Soir=" & arr(i, 14)
            rapport = rapport & " Nuit=" & arr(i, 15)
            rapport = rapport & vbCrLf
            
            ' Stocker dans dict
            If Not d.Exists(code) Then
                Dim v(1 To 4) As Double
                If IsNumeric(arr(i, 12)) Then v(1) = CDbl(arr(i, 12))
                If IsNumeric(arr(i, 13)) Then v(2) = CDbl(arr(i, 13))
                If IsNumeric(arr(i, 14)) Then v(3) = CDbl(arr(i, 14))
                If IsNumeric(arr(i, 15)) Then v(4) = CDbl(arr(i, 15))
                d.Add code, v
            End If
        End If
    Next i
    
    rapport = rapport & vbCrLf & "=== TEST LOOKUP ===" & vbCrLf
    
    ' Tester les lookups
    Dim testCodes As Variant
    testCodes = Array("C 15", "C15", "C 19", "C19", "C 20", "C20", "C 20 E", "C20E", "C 19 DI", "C19DI")
    
    Dim t As Long
    For t = 0 To UBound(testCodes)
        Dim testCode As String: testCode = testCodes(t)
        If d.Exists(testCode) Then
            rapport = rapport & "  [" & testCode & "] => TROUVE" & vbCrLf
        Else
            ' Try normalized
            Dim normCode As String: normCode = Replace(testCode, " ", "")
            If d.Exists(normCode) Then
                rapport = rapport & "  [" & testCode & "] => PAS TROUVE, mais [" & normCode & "] EXISTE" & vbCrLf
            Else
                rapport = rapport & "  [" & testCode & "] => PAS TROUVE" & vbCrLf
            End If
        End If
    Next t
    
    Debug.Print rapport
    MsgBox "Rapport dans Ctrl+G", vbInformation
End Sub

