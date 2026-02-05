' Debug pour vérifier le chargement des codes
Sub Debug_ChargementCodes()
    Dim wsCodesSpec As Worksheet
    Dim wsConfigCodes As Worksheet
    Dim wsConfig As Worksheet
    Dim dictCodes As Object
    Dim configGlobal As Object
    
    On Error Resume Next
    Set wsCodesSpec = ThisWorkbook.Sheets("Codes_Speciaux")
    Set wsConfigCodes = ThisWorkbook.Sheets("Config_Codes")
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    ' Charger config
    Set configGlobal = CreateObject("Scripting.Dictionary")
    If Not wsConfig Is Nothing Then
        Dim lr As Long, arr As Variant, i As Long
        lr = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
        If lr >= 2 Then
            arr = wsConfig.Range("A2:B" & lr).Value
            For i = 1 To UBound(arr, 1)
                If Trim(CStr(arr(i, 1))) <> "" Then
                    configGlobal(Trim(CStr(arr(i, 1)))) = Trim(CStr(arr(i, 2)))
                End If
            Next i
        End If
    End If
    
    ' Créer dictionnaire
    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    
    Dim msg As String
    msg = "=== CHARGEMENT CODES ===" & vbLf & vbLf
    
    ' 1. Charger Codes_Speciaux
    If wsCodesSpec Is Nothing Then
        msg = msg & "Codes_Speciaux: NON TROUVE" & vbLf
    Else
        Dim nbSpec As Long: nbSpec = dictCodes.Count
        ChargerSpeciauxDBG wsCodesSpec, dictCodes
        msg = msg & "Codes_Speciaux: " & (dictCodes.Count - nbSpec) & " charges" & vbLf
    End If
    
    ' 2. Charger Config_Codes
    If wsConfigCodes Is Nothing Then
        msg = msg & "Config_Codes: NON TROUVE" & vbLf
    Else
        Dim nbBefore As Long: nbBefore = dictCodes.Count
        ChargerConfigCodesDBG wsConfigCodes, dictCodes
        msg = msg & "Config_Codes: " & (dictCodes.Count - nbBefore) & " ajoutes" & vbLf
    End If
    
    msg = msg & vbLf & "TOTAL: " & dictCodes.Count & " codes" & vbLf & vbLf
    
    ' 3. Tester quelques codes clés
    msg = msg & "=== TEST CODES ===" & vbLf
    Dim testCodes As Variant
    testCodes = Array("8:30 16:30", "7 15:30", "7 13", "C 20 E", "C 19", "WE")
    
    Dim tc As Variant, vals As Variant
    For Each tc In testCodes
        If dictCodes.Exists(CStr(tc)) Then
            vals = dictCodes(CStr(tc))
            msg = msg & tc & ": M=" & vals(1) & " AM=" & vals(2)
            msg = msg & " S=" & vals(3) & " N=" & vals(4) & vbLf
        Else
            msg = msg & tc & ": NON TROUVE" & vbLf
        End If
    Next tc
    
    MsgBox msg, vbInformation, "Debug Chargement"
End Sub

Private Sub ChargerSpeciauxDBG(ws As Worksheet, d As Object)
    Dim lr As Long, arr As Variant, i As Long, k As Long
    Dim code As String
    Dim vals(1 To 11) As Double
    
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub
    arr = ws.Range("A2:F" & lr).Value
    
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" And Not d.Exists(code) Then
            For k = 1 To 11: vals(k) = 0: Next k
            vals(1) = NumValDBG(arr(i, 2))
            vals(2) = NumValDBG(arr(i, 3))
            vals(3) = NumValDBG(arr(i, 4))
            vals(4) = NumValDBG(arr(i, 5))
            d.Add code, vals
        End If
    Next i
End Sub

Private Sub ChargerConfigCodesDBG(ws As Worksheet, d As Object)
    Dim lr As Long, arr As Variant, i As Long, k As Long
    Dim code As String
    Dim v(1 To 11) As Double
    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
    
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub
    arr = ws.Range("A2:A" & lr).Value
    
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" And Not d.Exists(code) Then
            For k = 1 To 11: v(k) = 0: Next k
            
            If ParseCodeDBG(code, h1, f1, h2, f2) Then
                CalcPeriodesDBG h1, f1, h2, f2, v(1), v(2), v(3), v(4)
            End If
            
            d.Add code, v
        End If
    Next i
End Sub

Private Function NumValDBG(x As Variant) As Double
    If IsNumeric(x) Then NumValDBG = CDbl(x) Else NumValDBG = 0
End Function

Private Function HeureDecDBG(s As String) As Double
    Dim p() As String
    If InStr(s, ":") > 0 Then
        p = Split(s, ":")
        HeureDecDBG = CDbl(p(0)) + CDbl(p(1)) / 60
    ElseIf IsNumeric(s) Then
        HeureDecDBG = CDbl(s)
    Else
        HeureDecDBG = 0
    End If
End Function

Private Function ParseCodeDBG(c As String, _
    ByRef s1 As Double, ByRef e1 As Double, _
    ByRef s2 As Double, ByRef e2 As Double) As Boolean
    
    s1 = 0: e1 = 0: s2 = 0: e2 = 0
    Dim p() As String, tmp As String
    tmp = Trim(Replace(Replace(c, vbLf, " "), vbCr, " "))
    Do While InStr(tmp, "  ") > 0: tmp = Replace(tmp, "  ", " "): Loop
    p = Split(tmp, " ")
    
    On Error GoTo Err1
    If UBound(p) = 1 Then
        s1 = HeureDecDBG(p(0)): e1 = HeureDecDBG(p(1))
        ParseCodeDBG = True
    ElseIf UBound(p) >= 3 Then
        s1 = HeureDecDBG(p(0)): e1 = HeureDecDBG(p(1))
        s2 = HeureDecDBG(p(2)): e2 = HeureDecDBG(p(3))
        ParseCodeDBG = True
    End If
    Exit Function
Err1:
    ParseCodeDBG = False
End Function

Private Sub CalcPeriodesDBG(h1 As Double, f1 As Double, _
    h2 As Double, f2 As Double, _
    ByRef mat As Double, ByRef am As Double, _
    ByRef soi As Double, ByRef nui As Double)
    
    mat = 0: am = 0: soi = 0: nui = 0
    Dim heureMatin As Double, heureAM As Double
    
    ' Matin = 8h-12h
    heureMatin = OverlapDBG(h1, f1, 8, 12) + OverlapDBG(h2, f2, 8, 12)
    If heureMatin >= 4 Then
        mat = 1
    ElseIf heureMatin >= 2 Then
        mat = 0.5
    ElseIf heureMatin > 0 Then
        mat = Round(heureMatin / 4, 2)
    End If
    
    ' AM = 12h-16h30
    heureAM = OverlapDBG(h1, f1, 12, 16.5) + OverlapDBG(h2, f2, 12, 16.5)
    If heureAM >= 4 Then
        am = 1
    ElseIf heureAM >= 2 Then
        am = 0.5
    ElseIf heureAM > 0 Then
        am = Round(heureAM / 4.5, 2)
    End If
    
    ' Soir = après 16h30
    If f1 > 16.5 Or f2 > 16.5 Then soi = 1
    
    ' Nuit
    If h1 >= 19.5 Or f1 <= 7.25 Then nui = 1
End Sub

Private Function OverlapDBG(hd As Double, hf As Double, _
    td As Double, tf As Double) As Double
    
    Dim os As Double, oe As Double
    os = Application.Max(hd, td)
    oe = Application.Min(hf, tf)
    If oe > os Then OverlapDBG = oe - os Else OverlapDBG = 0
End Function
