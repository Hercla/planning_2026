' Module de DEBUG détaillé par jour
' Affiche qui est compté pour chaque période d'un jour donné

Sub Debug_JourDetail()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim wsCodesSpec As Worksheet, wsConfigCodes As Worksheet, wsConfig As Worksheet
    Dim dictCodes As Object, dictFonctions As Object, configGlobal As Object
    
    On Error Resume Next
    Set wsCodesSpec = ThisWorkbook.Sheets("Codes_Speciaux")
    Set wsConfigCodes = ThisWorkbook.Sheets("Config_Codes")
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    ' Charger config
    Set configGlobal = ChargerConfigDebug(wsConfig)
    
    ' Charger codes
    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    If Not wsCodesSpec Is Nothing Then ChargerSpeciauxDebug wsCodesSpec, dictCodes
    If Not wsConfigCodes Is Nothing Then ChargerConfigCodesDebug wsConfigCodes, dictCodes, configGlobal
    
    ' Charger fonctions
    Set dictFonctions = ChargerFonctionsDebug2()
    
    ' Demander le jour à analyser
    Dim colJour As Long
    colJour = Application.InputBox("Entrez le numéro de COLONNE du jour à analyser (ex: 3 pour colonne C = jour 2):", _
                                    "Debug Jour", 3, Type:=1)
    If colJour = 0 Then Exit Sub
    
    ' Paramètres
    Dim ligneDebut As Long: ligneDebut = 6
    Dim ligneFin As Long: ligneFin = 28
    Dim couleurIgnore As Long: couleurIgnore = 15849925
    
    If configGlobal.Exists("CHK_FirstPersonnelRow") Then ligneDebut = CLng(configGlobal("CHK_FirstPersonnelRow"))
    If configGlobal.Exists("CHK_IgnoreColor") Then couleurIgnore = CLng(configGlobal("CHK_IgnoreColor"))
    
    ' Analyser
    Dim i As Long, j As Long
    Dim cell As Range, code As String, nomPersonne As String
    Dim vals As Variant
    Dim estINF As Boolean
    Dim msg As String
    Dim tot(1 To 4) As Double, totINF(1 To 4) As Double
    Dim detailMatin As String, detailPM As String, detailSoir As String, detailNuit As String
    
    detailMatin = "": detailPM = "": detailSoir = "": detailNuit = ""
    
    For i = ligneDebut To ligneFin
        Set cell = ws.Cells(i, colJour)
        If cell.Interior.Color <> couleurIgnore Then
            code = Trim(CStr(cell.Value))
            nomPersonne = Trim(CStr(ws.Cells(i, 1).Value))
            
            If code <> "" And dictCodes.Exists(code) Then
                vals = dictCodes(code)
                
                ' Vérifier INF
                estINF = False
                If dictFonctions.Exists(nomPersonne) Then
                    If UCase(dictFonctions(nomPersonne)) = "INF" Then estINF = True
                End If
                
                ' Période 1 = Matin
                If vals(1) > 0 Then
                    tot(1) = tot(1) + vals(1)
                    If estINF Then
                        totINF(1) = totINF(1) + vals(1)
                        detailMatin = detailMatin & "  [INF] " & nomPersonne & " (" & code & ")" & vbLf
                    Else
                        detailMatin = detailMatin & "  " & nomPersonne & " (" & code & ")" & vbLf
                    End If
                End If
                
                ' Période 2 = Après-midi
                If vals(2) > 0 Then
                    tot(2) = tot(2) + vals(2)
                    If estINF Then
                        totINF(2) = totINF(2) + vals(2)
                        detailPM = detailPM & "  [INF] " & nomPersonne & " (" & code & ")" & vbLf
                    Else
                        detailPM = detailPM & "  " & nomPersonne & " (" & code & ")" & vbLf
                    End If
                End If
                
                ' Période 3 = Soir
                If vals(3) > 0 Then
                    tot(3) = tot(3) + vals(3)
                    If estINF Then
                        totINF(3) = totINF(3) + vals(3)
                        detailSoir = detailSoir & "  [INF] " & nomPersonne & " (" & code & ")" & vbLf
                    Else
                        detailSoir = detailSoir & "  " & nomPersonne & " (" & code & ")" & vbLf
                    End If
                End If
                
                ' Période 4 = Nuit
                If vals(4) > 0 Then
                    tot(4) = tot(4) + vals(4)
                    If estINF Then
                        totINF(4) = totINF(4) + vals(4)
                        detailNuit = detailNuit & "  [INF] " & nomPersonne & " (" & code & ")" & vbLf
                    Else
                        detailNuit = detailNuit & "  " & nomPersonne & " (" & code & ")" & vbLf
                    End If
                End If
            End If
        End If
    Next i
    
    ' Afficher résultats
    msg = "=== JOUR COLONNE " & colJour & " ===" & vbLf & vbLf
    
    msg = msg & "MATIN: " & tot(1) & " (" & totINF(1) & " INF)" & vbLf
    If detailMatin <> "" Then msg = msg & detailMatin
    msg = msg & vbLf
    
    msg = msg & "APRES-MIDI: " & tot(2) & " (" & totINF(2) & " INF)" & vbLf
    If detailPM <> "" Then msg = msg & detailPM
    msg = msg & vbLf
    
    msg = msg & "SOIR: " & tot(3) & " (" & totINF(3) & " INF)" & vbLf
    If detailSoir <> "" Then msg = msg & detailSoir
    msg = msg & vbLf
    
    msg = msg & "NUIT: " & tot(4) & " (" & totINF(4) & " INF)" & vbLf
    If detailNuit <> "" Then msg = msg & detailNuit
    
    MsgBox msg, vbInformation, "Debug Detail Jour"
End Sub

' === FONCTIONS COPIEES DU MODULE PRINCIPAL ===

Private Function ChargerConfigDebug(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If ws Is Nothing Then Set ChargerConfigDebug = d: Exit Function
    
    Dim lr As Long, arr As Variant, i As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Set ChargerConfigDebug = d: Exit Function
    
    arr = ws.Range("A2:B" & lr).Value
    For i = 1 To UBound(arr, 1)
        If Trim(CStr(arr(i, 1))) <> "" Then
            d(Trim(CStr(arr(i, 1)))) = Trim(CStr(arr(i, 2)))
        End If
    Next i
    Set ChargerConfigDebug = d
End Function

Private Sub ChargerSpeciauxDebug(ws As Worksheet, d As Object)
    Dim lr As Long, arr As Variant, i As Long
    Dim code As String, vals(1 To 11) As Double
    
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub
    
    arr = ws.Range("A2:L" & lr).Value
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" Then
            vals(1) = Val(arr(i, 2))
            vals(2) = Val(arr(i, 3))
            vals(3) = Val(arr(i, 4))
            vals(4) = Val(arr(i, 5))
            vals(5) = Val(arr(i, 6))
            vals(6) = Val(arr(i, 7))
            vals(7) = Val(arr(i, 8))
            vals(8) = Val(arr(i, 9))
            vals(9) = Val(arr(i, 10))
            vals(10) = Val(arr(i, 11))
            vals(11) = Val(arr(i, 12))
            d(code) = vals
        End If
    Next i
End Sub

Private Sub ChargerConfigCodesDebug(ws As Worksheet, d As Object, cfg As Object)
    Dim lr As Long, arr As Variant, i As Long
    Dim code As String, vals(1 To 11) As Double
    Dim colCode As Long, colMatin As Long, colPM As Long, colSoir As Long, colNuit As Long
    
    colCode = 1: colMatin = 2: colPM = 3: colSoir = 4: colNuit = 5
    If cfg.Exists("CFGCODES_Col_Code") Then colCode = CLng(cfg("CFGCODES_Col_Code"))
    If cfg.Exists("CFGCODES_Col_Matin") Then colMatin = CLng(cfg("CFGCODES_Col_Matin"))
    If cfg.Exists("CFGCODES_Col_PM") Then colPM = CLng(cfg("CFGCODES_Col_PM"))
    If cfg.Exists("CFGCODES_Col_Soir") Then colSoir = CLng(cfg("CFGCODES_Col_Soir"))
    If cfg.Exists("CFGCODES_Col_Nuit") Then colNuit = CLng(cfg("CFGCODES_Col_Nuit"))
    
    lr = ws.Cells(ws.Rows.Count, colCode).End(xlUp).Row
    If lr < 2 Then Exit Sub
    
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lr, 20)).Value
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, colCode)))
        If code <> "" And Not d.Exists(code) Then
            vals(1) = Val(arr(i, colMatin))
            vals(2) = Val(arr(i, colPM))
            vals(3) = Val(arr(i, colSoir))
            vals(4) = Val(arr(i, colNuit))
            vals(5) = 0: vals(6) = 0: vals(7) = 0: vals(8) = 0
            vals(9) = 0: vals(10) = 0: vals(11) = 0
            d(code) = vals
        End If
    Next i
End Sub

Private Function ChargerFonctionsDebug2() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    Dim wsPersonnel As Worksheet
    On Error Resume Next
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If wsPersonnel Is Nothing Then
        Set ChargerFonctionsDebug2 = d
        Exit Function
    End If
    
    Dim lr As Long, arr As Variant, i As Long
    Dim nom As String, prenom As String, cleNomPrenom As String, fonction As String
    
    lr = wsPersonnel.Cells(wsPersonnel.Rows.Count, "B").End(xlUp).Row
    If lr < 2 Then Set ChargerFonctionsDebug2 = d: Exit Function
    
    arr = wsPersonnel.Range("B2:E" & lr).Value
    For i = 1 To UBound(arr, 1)
        nom = Trim(CStr(arr(i, 1)))
        prenom = Trim(CStr(arr(i, 2)))
        fonction = Trim(CStr(arr(i, 4)))
        cleNomPrenom = nom & "_" & prenom
        
        If cleNomPrenom <> "_" And Not d.Exists(cleNomPrenom) Then
            d.Add cleNomPrenom, fonction
        End If
    Next i
    
    Set ChargerFonctionsDebug2 = d
End Function
