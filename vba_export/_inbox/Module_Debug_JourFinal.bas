' Debug FINAL avec logique binaire corrigée
Sub Debug_JourFinal()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim wsCodesSpec As Worksheet, wsConfigCodes As Worksheet, wsConfig As Worksheet
    Dim dictCodes As Object, dictFonctions As Object, configGlobal As Object
    
    On Error Resume Next
    Set wsCodesSpec = ThisWorkbook.Sheets("Codes_Speciaux")
    Set wsConfigCodes = ThisWorkbook.Sheets("Config_Codes")
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    ' Charger config
    Set configGlobal = CreateObject("Scripting.Dictionary")
    
    ' Charger codes
    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    If Not wsCodesSpec Is Nothing Then ChargerSpecFinal wsCodesSpec, dictCodes
    If Not wsConfigCodes Is Nothing Then ChargerCfgFinal wsConfigCodes, dictCodes
    
    ' Charger fonctions
    Set dictFonctions = ChargerFonctionsFinal()
    
    ' Demander la colonne
    Dim colJour As Long
    colJour = Application.InputBox("Colonne du jour (ex: 4 pour Mar 3):", _
                                    "Debug", 4, Type:=1)
    If colJour = 0 Then Exit Sub
    
    ' Paramètres
    Dim ligneDebut As Long: ligneDebut = 6
    Dim ligneFin As Long: ligneFin = 28
    Dim couleurIgnore As Long: couleurIgnore = 15849925
    
    ' Calculer
    Dim i As Long, j As Long
    Dim cell As Range, code As String, nomPersonne As String
    Dim tot(1 To 4) As Double, totINF(1 To 4) As Double
    Dim vals As Variant, estINF As Boolean
    Dim detMatin As String, detPM As String
    
    For j = 1 To 4: tot(j) = 0: totINF(j) = 0: Next j
    detMatin = "": detPM = ""
    
    For i = ligneDebut To ligneFin
        Set cell = ws.Cells(i, colJour)
        If cell.Interior.Color <> couleurIgnore Then
            code = Trim(CStr(cell.Value))
            If code <> "" And dictCodes.Exists(code) Then
                vals = dictCodes(code)
                
                For j = 1 To 4: tot(j) = tot(j) + vals(j): Next j
                
                nomPersonne = Trim(CStr(ws.Cells(i, 1).Value))
                estINF = False
                If dictFonctions.Exists(nomPersonne) Then
                    If UCase(dictFonctions(nomPersonne)) = "INF" Then estINF = True
                End If
                
                If estINF Then
                    For j = 1 To 4: totINF(j) = totINF(j) + vals(j): Next j
                End If
                
                If vals(1) > 0 Then
                    If estINF Then detMatin = detMatin & "[INF] "
                    detMatin = detMatin & nomPersonne & " (" & code & ")" & vbLf
                End If
                If vals(2) > 0 Then
                    If estINF Then detPM = detPM & "[INF] "
                    detPM = detPM & nomPersonne & " (" & code & ")" & vbLf
                End If
            End If
        End If
    Next i
    
    Dim msg As String
    msg = "=== COL " & colJour & " ===" & vbLf & vbLf
    msg = msg & "MATIN: " & tot(1) & " (" & totINF(1) & " INF)" & vbLf
    If detMatin <> "" Then msg = msg & detMatin
    msg = msg & vbLf
    msg = msg & "APRES-MIDI: " & tot(2) & " (" & totINF(2) & " INF)" & vbLf
    If detPM <> "" Then msg = msg & detPM
    msg = msg & vbLf
    msg = msg & "SOIR: " & tot(3) & " (" & totINF(3) & " INF)" & vbLf
    msg = msg & "NUIT: " & tot(4) & " (" & totINF(4) & " INF)"
    
    MsgBox msg, vbInformation, "Debug Final"
End Sub

Private Sub ChargerSpecFinal(ws As Worksheet, d As Object)
    Dim lr As Long, arr As Variant, i As Long, k As Long
    Dim code As String, vals(1 To 11) As Double
    
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub
    arr = ws.Range("A2:F" & lr).Value
    
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" And Not d.Exists(code) Then
            For k = 1 To 11: vals(k) = 0: Next k
            vals(1) = NumV(arr(i, 2))
            vals(2) = NumV(arr(i, 3))
            vals(3) = NumV(arr(i, 4))
            vals(4) = NumV(arr(i, 5))
            d.Add code, vals
        End If
    Next i
End Sub

Private Sub ChargerCfgFinal(ws As Worksheet, d As Object)
    Dim lr As Long, arr As Variant, i As Long, k As Long
    Dim code As String, v(1 To 11) As Double
    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
    
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub
    arr = ws.Range("A2:A" & lr).Value
    
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" And Not d.Exists(code) Then
            For k = 1 To 11: v(k) = 0: Next k
            If ParseCd(code, h1, f1, h2, f2) Then
                ' LOGIQUE BINAIRE
                ' Matin = 1 si commence avant 13h
                If h1 < 13 Or h2 < 13 Then v(1) = 1
                ' AM = 1 si finit après 13h
                If f1 > 13 Or f2 > 13 Then v(2) = 1
                ' Soir = 1 si finit après 16h30
                If f1 > 16.5 Or f2 > 16.5 Then v(3) = 1
                ' Nuit = 1 si commence après 19h30 ou finit avant 7h15
                If h1 >= 19.5 Or (f1 > 0 And f1 <= 7.25) Then v(4) = 1
            End If
            d.Add code, v
        End If
    Next i
End Sub

Private Function ChargerFonctionsFinal() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim wsP As Worksheet
    On Error Resume Next
    Set wsP = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    If wsP Is Nothing Then Set ChargerFonctionsFinal = d: Exit Function
    
    Dim lr As Long, arr As Variant, i As Long
    lr = wsP.Cells(wsP.Rows.Count, "B").End(xlUp).Row
    If lr < 2 Then Set ChargerFonctionsFinal = d: Exit Function
    
    arr = wsP.Range("B2:E" & lr).Value
    For i = 1 To UBound(arr, 1)
        Dim nom As String, prenom As String, cle As String, fn As String
        nom = Trim(CStr(arr(i, 1)))
        prenom = Trim(CStr(arr(i, 2)))
        fn = Trim(CStr(arr(i, 4)))
        cle = nom & "_" & prenom
        If cle <> "_" And Not d.Exists(cle) Then d.Add cle, fn
    Next i
    Set ChargerFonctionsFinal = d
End Function

Private Function NumV(x As Variant) As Double
    If IsNumeric(x) Then NumV = CDbl(x) Else NumV = 0
End Function

Private Function HeureDc(s As String) As Double
    Dim p() As String
    If InStr(s, ":") > 0 Then
        p = Split(s, ":")
        HeureDc = CDbl(p(0)) + CDbl(p(1)) / 60
    ElseIf IsNumeric(s) Then
        HeureDc = CDbl(s)
    Else
        HeureDc = 0
    End If
End Function

Private Function ParseCd(c As String, _
    ByRef s1 As Double, ByRef e1 As Double, _
    ByRef s2 As Double, ByRef e2 As Double) As Boolean
    s1 = 0: e1 = 0: s2 = 0: e2 = 0
    Dim p() As String, tmp As String
    tmp = Trim(Replace(Replace(c, vbLf, " "), vbCr, " "))
    Do While InStr(tmp, "  ") > 0: tmp = Replace(tmp, "  ", " "): Loop
    p = Split(tmp, " ")
    On Error GoTo Err1
    If UBound(p) = 1 Then
        s1 = HeureDc(p(0)): e1 = HeureDc(p(1))
        ParseCd = True
    ElseIf UBound(p) >= 3 Then
        s1 = HeureDc(p(0)): e1 = HeureDc(p(1))
        s2 = HeureDc(p(2)): e2 = HeureDc(p(3))
        ParseCd = True
    End If
    Exit Function
Err1:
    ParseCd = False
End Function
