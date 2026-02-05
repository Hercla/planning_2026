Attribute VB_Name = "Module_Planning_Core"
'Attribute VB_Name = "Module_Planning_Core"
Option Explicit

' =========================================================================
' MODULE: Module_Planning_Core
' BUT:    Fonctions communes partag�es entre Calculer_Totaux et Replacements
' DATE:   Janvier 2026
' =========================================================================

' =============================================================================
' SECTION 1: GESTION DE LA CONFIGURATION
' =============================================================================

Public Function ChargerConfig(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If ws Is Nothing Then Set ChargerConfig = d: Exit Function
    
    Dim lr As Long, arr As Variant, i As Long
    Dim cle As String, valeur As Variant
    lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Set ChargerConfig = d: Exit Function
    
    arr = ws.Range("A2:B" & lr).value
    For i = 1 To UBound(arr, 1)
        cle = Trim(CStr(arr(i, 1)))
        valeur = arr(i, 2)
        
        If cle <> "" And Not d.Exists(cle) Then
            d(cle) = valeur
        End If
    Next i
    Set ChargerConfig = d
End Function

Public Function CfgLongFromDict(d As Object, k As String, def As Long) As Long
    If d.Exists(k) Then
        If IsNumeric(d(k)) Then CfgLongFromDict = CLng(d(k)): Exit Function
    End If
    CfgLongFromDict = def
End Function

Public Function CfgStr(d As Object, k As String) As String
    If d.Exists(k) Then CfgStr = CStr(d(k)) Else CfgStr = ""
End Function

' =============================================================================
' SECTION 2: CHARGEMENT DES CODES HORAIRES
' =============================================================================

Public Sub ChargerSpeciaux(ws As Worksheet, d As Object)
    Dim lr As Long, arr As Variant, i As Long, k As Long
    Dim code As String, cc As String
    Dim v(1 To 11) As Double
    
    lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Exit Sub
    arr = ws.Range("A2:G" & lr).value
    
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" And Not d.Exists(code) Then
            For k = 1 To 11: v(k) = 0: Next k
            
            v(5) = NumVal(arr(i, 2))
            v(6) = NumVal(arr(i, 3))
            v(1) = NumVal(arr(i, 4))
            v(2) = NumVal(arr(i, 5))
            v(3) = NumVal(arr(i, 6))
            v(4) = NumVal(arr(i, 7))
            
            cc = UCase(Replace(code, " ", ""))
            If cc Like "C19*" Then v(11) = 1: If v(6) = 0 Then v(6) = 1
            If cc Like "C20E*" Then v(10) = 1
            If cc Like "C20*" And Not cc Like "C20E*" Then v(9) = 1
            If cc Like "C15*" Then v(8) = 1
            
            d.Add code, v
        End If
    Next i
End Sub

Public Sub ChargerConfigCodes(ws As Worksheet, d As Object, cfg As Object)
    Dim lr As Long, arr As Variant, i As Long, k As Long
    Dim code As String
    Dim v(1 To 11) As Double
    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
    Dim hasExtendedCols As Boolean
    Dim hStart As String, hPauseS As String, hPauseE As String, hEnd As String
    Dim manF645 As Variant, manF78 As Variant
    Dim manMatin As Variant, manPM As Variant, manSoir As Variant, manNuit As Variant
    
    Dim sC15 As String: sC15 = CfgStr(cfg, "SPECIAL_C15")
    Dim sC20 As String: sC20 = CfgStr(cfg, "SPECIAL_C20")
    Dim sC20E As String: sC20E = CfgStr(cfg, "SPECIAL_C20E")
    Dim sC19 As String: sC19 = CfgStr(cfg, "SPECIAL_C19")
    
    lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Exit Sub
    
    hasExtendedCols = (ws.Cells(1, 6).value <> "")
    
    If hasExtendedCols Then
        arr = ws.Range("A2:O" & lr).value
    Else
        arr = ws.Range("A2:A" & lr).value
    End If
    
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" And Not d.Exists(code) Then
            For k = 1 To 11: v(k) = 0: Next k
            h1 = 0: f1 = 0: h2 = 0: f2 = 0
            
            If hasExtendedCols Then
                manF645 = arr(i, 10)
                manF78 = arr(i, 11)
                manMatin = arr(i, 12)
                manPM = arr(i, 13)
                manSoir = arr(i, 14)
                manNuit = arr(i, 15)
                
                If IsNumeric(manMatin) And manMatin <> "" Then v(1) = CDbl(manMatin)
                If IsNumeric(manPM) And manPM <> "" Then v(2) = CDbl(manPM)
                If IsNumeric(manSoir) And manSoir <> "" Then v(3) = CDbl(manSoir)
                If IsNumeric(manNuit) And manNuit <> "" Then v(4) = CDbl(manNuit)
                
                If IsNumeric(manF645) And manF645 <> "" Then v(5) = CDbl(manF645)
                If IsNumeric(manF78) And manF78 <> "" Then v(6) = CDbl(manF78)
                
                hStart = Trim(CStr(arr(i, 6)))
                hPauseS = Trim(CStr(arr(i, 7)))
                hPauseE = Trim(CStr(arr(i, 8)))
                hEnd = Trim(CStr(arr(i, 9)))
                
                If hStart <> "" And hEnd <> "" Then
                    h1 = HeureDec(hStart)
                    If hPauseS <> "" And hPauseE <> "" Then
                        f1 = HeureDec(hPauseS)
                        h2 = HeureDec(hPauseE)
                        f2 = HeureDec(hEnd)
                    Else
                        f1 = HeureDec(hEnd)
                    End If
                    
                    Dim autoMat As Double, autoPM As Double
                    Dim autoSoir As Double, autoNuit As Double
                    CalcPeriodes h1, f1, h2, f2, autoMat, autoPM, autoSoir, autoNuit
                    
                    Dim uc As String: uc = UCase(code)
                    If Left(uc, 1) = "C" Or Left(uc, 4) = "SA C" Or Left(uc, 4) = "DI C" Then
                        autoPM = 0
                    End If
                    
                    If uc = "C 19" Or uc = "C 19 SA" Or uc = "C 19 DI" Then
                        If autoMat = 0 Then autoMat = 1
                        If autoSoir = 0 Then autoSoir = 1
                    End If
                    
                    Dim cc As String: cc = Replace(uc, " ", "")
                    Dim isC20EName As Boolean, isC20Name As Boolean, isC15Name As Boolean
                    isC20EName = (cc Like "C20E*")
                    isC20Name = (cc Like "C20*") And Not isC20EName
                    isC15Name = (cc Like "C15*")
                    
                    If cc Like "C19*" Then v(11) = 1: If v(6) = 0 Then v(6) = 1
                    If IsCodeC19Pattern(h1, f1, h2, f2) Then v(11) = 1
                    
                    If isC20EName Then v(10) = 1
                    If IsCodeC20E(h1, f1, h2, f2) Then v(10) = 1
                    
                    If isC20Name Then v(9) = 1
                    If IsCodeC20Pattern(h1, f1, h2, f2) Then v(9) = 1
                    
                    If isC15Name Then v(8) = 1
                    If IsCodeC15Pattern(h1, f1, h2, f2) Then v(8) = 1
                    
                    If Not (IsNumeric(manMatin) And manMatin <> "") Then v(1) = autoMat
                    If Not (IsNumeric(manPM) And manPM <> "") Then v(2) = autoPM
                    If Not (IsNumeric(manSoir) And manSoir <> "") Then v(3) = autoSoir
                    If Not (IsNumeric(manNuit) And manNuit <> "") Then v(4) = autoNuit
                    
                    Dim autoF645 As Double, autoF78 As Double, autoP81630 As Double
                    CalcPresSpec h1, f1, h2, f2, autoF645, autoF78, autoP81630
                    
                    If Not (IsNumeric(manF645) And manF645 <> "") Then v(5) = autoF645
                    If Not (IsNumeric(manF78) And manF78 <> "") Then v(6) = autoF78
                    v(7) = autoP81630
                    
                    If MatchSpecial(h1, f1, h2, f2, sC15) Then v(8) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC20) Then v(9) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC20E) Then v(10) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC19) Then v(11) = 1

                    If isC20EName Then
                        v(10) = 1: v(9) = 0: v(8) = 0
                    ElseIf isC20Name Then
                        v(9) = 1: v(10) = 0: v(8) = 0
                    ElseIf isC15Name Then
                        v(8) = 1: v(9) = 0: v(10) = 0
                    Else
                        If v(8) > 0 Then v(9) = 0: v(10) = 0
                        If v(9) > 0 Or v(10) > 0 Then v(8) = 0
                    End If
                End If
            End If
            
            If v(1) = 0 And v(2) = 0 And v(3) = 0 And v(4) = 0 Then
                If ParseCode(code, h1, f1, h2, f2) Then
                    CalcPeriodes h1, f1, h2, f2, v(1), v(2), v(3), v(4)
                    CalcPresSpec h1, f1, h2, f2, v(5), v(6), v(7)
                    cc = Replace(UCase(code), " ", "")
                    isC20EName = (cc Like "C20E*")
                    isC20Name = (cc Like "C20*") And Not isC20EName
                    isC15Name = (cc Like "C15*")
                    If MatchSpecial(h1, f1, h2, f2, sC15) Then v(8) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC20) Then v(9) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC20E) Then v(10) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC19) Then v(11) = 1
                    If isC20EName Then
                        v(10) = 1: v(9) = 0: v(8) = 0
                    ElseIf isC20Name Then
                        v(9) = 1: v(10) = 0: v(8) = 0
                    ElseIf isC15Name Then
                        v(8) = 1: v(9) = 0: v(10) = 0
                    Else
                        If v(8) > 0 Then v(9) = 0: v(10) = 0
                        If v(9) > 0 Or v(10) > 0 Then v(8) = 0
                    End If
                End If
            End If
            
            d.Add code, v
        End If
    Next i
End Sub

' =============================================================================
' SECTION 3: PARSING DES CODES HORAIRES
' =============================================================================

Public Function HeureDec(s As Variant) As Double
    Dim p() As String
    Dim strVal As String
    
    If IsEmpty(s) Or s = "" Then
        HeureDec = 0
        Exit Function
    End If
    
    If IsNumeric(s) Then
        If CDbl(s) < 1 And CDbl(s) > 0 Then
            HeureDec = CDbl(s) * 24
            Exit Function
        ElseIf CDbl(s) >= 1 And CDbl(s) < 25 Then
            HeureDec = CDbl(s)
            Exit Function
        End If
    End If
    
    strVal = CStr(s)
    If InStr(strVal, ":") > 0 Then
        p = Split(strVal, ":")
        On Error Resume Next
        HeureDec = CDbl(p(0)) + CDbl(p(1)) / 60
        On Error GoTo 0
    ElseIf IsNumeric(strVal) Then
        HeureDec = CDbl(strVal)
    Else
        HeureDec = 0
    End If
End Function

Public Function ParseCode(c As String, ByRef s1 As Double, ByRef e1 As Double, ByRef s2 As Double, ByRef e2 As Double) As Boolean
    s1 = 0: e1 = 0: s2 = 0: e2 = 0
    Dim p() As String, tmp As String
    tmp = Trim(Replace(Replace(c, vbLf, " "), vbCr, " "))
    Do While InStr(tmp, "  ") > 0: tmp = Replace(tmp, "  ", " "): Loop
    p = Split(tmp, " ")
    
    On Error GoTo Err1
    If UBound(p) = 1 Then
        s1 = HeureDec(p(0)): e1 = HeureDec(p(1))
        ParseCode = True
    ElseIf UBound(p) >= 3 Then
        s1 = HeureDec(p(0)): e1 = HeureDec(p(1))
        s2 = HeureDec(p(2)): e2 = HeureDec(p(3))
        ParseCode = True
    End If
    Exit Function
Err1:
    ParseCode = False
End Function

' =============================================================================
' SECTION 4: CALCUL DES P�RIODES ET PR�SENCES
' =============================================================================

Public Sub CalcPeriodes(h1 As Double, f1 As Double, h2 As Double, f2 As Double, ByRef mat As Double, ByRef am As Double, ByRef soi As Double, ByRef nui As Double)
    mat = 0: am = 0: soi = 0: nui = 0
    
    Dim fin As Double
    fin = IIf(f2 > 0, f2, f1)
    
    If h1 = 0 And f1 = 0 Then Exit Sub
    
    ' MATIN : Presence si debut < 13h
    If h1 < 13 Then
        mat = 1
    ElseIf h2 > 0 And h2 < 13 Then
        mat = 1
    End If
    
    ' APRES-MIDI (PM) : Presence si fin > 13h
    If fin > 13 Then
        am = 1
    End If
    If h2 > 0 And f2 > 13 Then
        am = 1
    End If
    
    ' SOIR : Presence si fin > 16h30
    If fin > 17.5 Then
        soi = 1
    ElseIf fin > 16.5 Then
        soi = 0.5
    End If
    
    ' NUIT
    If h1 >= 19.5 Or (fin > 0 And fin <= 7.25) Then
        If Abs(fin - 24) < 0.1 Or fin = 0 Then
            nui = 0.5
        Else
            nui = 1
        End If
    End If
End Sub

Public Sub CalcPresSpec(h1 As Double, f1 As Double, h2 As Double, f2 As Double, ByRef p645 As Double, ByRef p78 As Double, ByRef p81630 As Double)
    p645 = 0: p78 = 0: p81630 = 0
    If h1 <= 6.75 Then p645 = 1
    If h1 < 8 And f1 > 7 Then p78 = 1
    If Abs(f1 - 16.5) < 0.25 Or Abs(f2 - 16.5) < 0.25 Then p81630 = 1
End Sub

' =============================================================================
' SECTION 5: D�TECTION DES CODES SP�CIAUX (C15, C19, C20, C20E)
' =============================================================================

Public Function IsCodeC15(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC15 = False
    Dim finEffective As Double
    finEffective = IIf(f2 > 0, f2, f1)
    If finEffective >= 15 And finEffective <= 15.5 Then IsCodeC15 = True
End Function

Public Function IsCodeC20(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC20 = False
    Dim finEffective As Double
    finEffective = IIf(f2 > 0, f2, f1)
    If finEffective >= 19.75 And finEffective <= 20.25 Then IsCodeC20 = True
End Function

Public Function IsCodeC20E(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC20E = False
    Dim finEffective As Double
    finEffective = IIf(f2 > 0, f2, f1)
    If finEffective > 20.25 And finEffective <= 21 Then IsCodeC20E = True
End Function

Public Function IsCodeC19(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC19 = False
    Dim finEffective As Double
    finEffective = IIf(f2 > 0, f2, f1)
    If finEffective >= 18.75 And finEffective <= 19.25 Then IsCodeC19 = True
End Function

Public Function IsCodeC15Pattern(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC15Pattern = False
    If h2 = 0 Or f2 = 0 Then Exit Function
    
    If f2 >= 19.75 And f2 <= 20.75 Then
        If f1 >= 11.5 And f1 <= 13 And h2 >= 15.5 And h2 <= 17 Then
            IsCodeC15Pattern = True
        End If
    End If
End Function

Public Function IsCodeC19Pattern(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC19Pattern = False
    If h2 = 0 Or f2 = 0 Then Exit Function
    
    If f2 >= 18.75 And f2 <= 19.25 Then
        If h1 <= 8 And f1 >= 11 And f1 <= 12 Then
            IsCodeC19Pattern = True
        End If
    End If
End Function

Public Function IsCodeC20Pattern(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC20Pattern = False
    If h2 = 0 Or f2 = 0 Then Exit Function
    
    If f2 >= 19.75 And f2 <= 20.25 Then
        If f1 >= 11.5 And f1 <= 12.5 And h2 >= 15.5 And h2 <= 16.5 Then
            IsCodeC20Pattern = True
        End If
    End If
End Function

' =============================================================================
' SECTION 6: GESTION DES JOURS F�RI�S
' =============================================================================

Public Function BuildFeriesBE(ByVal annee As Long) As Object
    Dim feries As Object
    Set feries = CreateObject("Scripting.Dictionary")
    Dim paques As Date
    
    paques = CalculerPaques(annee)
    
    On Error Resume Next
    feries.Add CStr(DateSerial(annee, 1, 1)), True
    feries.Add CStr(paques + 1), True
    feries.Add CStr(DateSerial(annee, 5, 1)), True
    feries.Add CStr(paques + 39), True
    feries.Add CStr(paques + 50), True
    feries.Add CStr(DateSerial(annee, 7, 21)), True
    feries.Add CStr(DateSerial(annee, 8, 15)), True
    feries.Add CStr(DateSerial(annee, 11, 1)), True
    feries.Add CStr(DateSerial(annee, 11, 11)), True
    feries.Add CStr(DateSerial(annee, 12, 25)), True
    On Error GoTo 0
    
    Set BuildFeriesBE = feries
End Function

Public Function CalculerPaques(ByVal annee As Long) As Date
    Dim a As Integer, b As Integer, c As Integer
    Dim d As Integer, e As Integer, f As Integer
    Dim g As Integer, h As Integer, i As Integer
    Dim k As Integer, l As Integer, m As Integer
    Dim mois As Integer, jour As Integer
    
    a = annee Mod 19
    b = annee \ 100
    c = annee Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    mois = (h + l - 7 * m + 114) \ 31
    jour = ((h + l - 7 * m + 114) Mod 31) + 1
    
    CalculerPaques = DateSerial(annee, mois, jour)
End Function

Public Function EstDansFeries(ByVal d As Date, ByVal feries As Object) As Boolean
    EstDansFeries = feries.Exists(CStr(d))
End Function

Public Function DateFromMoisNom(ByVal nomMois As String, ByVal jour As Long, ByVal annee As Long) As Date
    Dim moisNum As Integer
    Select Case LCase(Left(nomMois, 4))
        Case "janv": moisNum = 1
        Case "fev", "fevr": moisNum = 2
        Case "mars": moisNum = 3
        Case "avri": moisNum = 4
        Case "mai": moisNum = 5
        Case "juin": moisNum = 6
        Case "juil": moisNum = 7
        Case "aout", "aout": moisNum = 8
        Case "sept": moisNum = 9
        Case "oct", "octo": moisNum = 10
        Case "nov", "nove": moisNum = 11
        Case "dec", "dece": moisNum = 12
        Case Else: moisNum = 1
    End Select
    DateFromMoisNom = DateSerial(annee, moisNum, jour)
End Function

' =============================================================================
' SECTION 7: CHARGEMENT PERSONNEL ET EXCLUSIONS
' =============================================================================

Public Function ChargerFonctionsPersonnel() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    Dim wsPersonnel As Worksheet
    On Error Resume Next
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If wsPersonnel Is Nothing Then
        Set ChargerFonctionsPersonnel = d
        Exit Function
    End If
    
    Dim lr As Long, arr As Variant, i As Long
    Dim nom As String, prenom As String, cleNomPrenom As String, fonction As String
    
    lr = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
    If lr < 2 Then Set ChargerFonctionsPersonnel = d: Exit Function
    
    arr = wsPersonnel.Range("B2:E" & lr).value
    For i = 1 To UBound(arr, 1)
        nom = Trim(CStr(arr(i, 1)))
        prenom = Trim(CStr(arr(i, 2)))
        fonction = Trim(CStr(arr(i, 4)))
        
        cleNomPrenom = nom & "_" & prenom
        
        If cleNomPrenom <> "_" And Not d.Exists(cleNomPrenom) Then
            d.Add cleNomPrenom, fonction
        End If
    Next i
    
    Set ChargerFonctionsPersonnel = d
End Function

Public Function ChargerCEFAEnFormation(nomMois As String, couleurFormation As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    Dim wsPersonnel As Worksheet
    On Error Resume Next
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If wsPersonnel Is Nothing Then
        Set ChargerCEFAEnFormation = d
        Exit Function
    End If
    
    Dim colMoisPct As Long
    Dim moisNum As Long: moisNum = MoisNumero(nomMois)
    colMoisPct = 29 + (moisNum - 1) * 2
    
    Dim lr As Long, i As Long
    Dim nom As String, prenom As String, cleNomPrenom As String, fonction As String
    Dim cellPct As Range
    
    lr = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
    If lr < 2 Then Set ChargerCEFAEnFormation = d: Exit Function
    
    For i = 2 To lr
        fonction = UCase(Trim(CStr(wsPersonnel.Cells(i, "E").value)))
        
        If fonction = "CEFA" Then
            nom = Trim(CStr(wsPersonnel.Cells(i, "B").value))
            prenom = Trim(CStr(wsPersonnel.Cells(i, "C").value))
            cleNomPrenom = nom & "_" & prenom
            
            Set cellPct = wsPersonnel.Cells(i, colMoisPct)
            If cellPct.Interior.Color = couleurFormation Then
                d(cleNomPrenom) = True
            End If
        End If
    Next i
    
    Set ChargerCEFAEnFormation = d
End Function

Public Function ChargerExclusionsCalcul() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config_Exceptions")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ChargerExclusionsCalcul = Empty
        Exit Function
    End If
    
    Dim lr As Long
    lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then
        ChargerExclusionsCalcul = Empty
        Exit Function
    End If
    
    ChargerExclusionsCalcul = ws.Range("A2:F" & lr).value
End Function

' =============================================================================
' SECTION 8: FONCTIONS UTILITAIRES
' =============================================================================

Public Function NumVal(x As Variant) As Double
    If IsNumeric(x) Then NumVal = CDbl(x) Else NumVal = 0
End Function

Public Function MatchSpecial(h1 As Double, f1 As Double, h2 As Double, f2 As Double, def As String) As Boolean
    MatchSpecial = False
    If def = "" Then Exit Function
    Dim p() As String: p = Split(def, " ")
    If UBound(p) < 3 Then Exit Function
    Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double
    c1 = HeureDec(p(0)): c2 = HeureDec(p(1)): c3 = HeureDec(p(2)): c4 = HeureDec(p(3))
    Const t As Double = 0.02
    If Abs(h1 - c1) < t And Abs(f1 - c2) < t And Abs(h2 - c3) < t And Abs(f2 - c4) < t Then MatchSpecial = True
End Function

Public Function MoisNumero(nomMois As String) As Long
    Select Case LCase(Left(nomMois, 4))
        Case "janv": MoisNumero = 1
        Case "fev", "fevr": MoisNumero = 2
        Case "mars": MoisNumero = 3
        Case "avri": MoisNumero = 4
        Case "mai": MoisNumero = 5
        Case "juin": MoisNumero = 6
        Case "juil": MoisNumero = 7
        Case "aout", "aout": MoisNumero = 8
        Case "sept": MoisNumero = 9
        Case "oct", "octo": MoisNumero = 10
        Case "nov", "nove": MoisNumero = 11
        Case "dec", "dece": MoisNumero = 12
        Case Else: MoisNumero = 1
    End Select
End Function

Public Function GetJourNom(numJour As Variant, annee As Long, nomMois As String) As String
    On Error GoTo Err1
    Dim d As Date
    Dim moisNum As Long: moisNum = MoisNumero(nomMois)
    d = DateSerial(annee, moisNum, CLng(numJour))
    
    Select Case Weekday(d, vbMonday)
        Case 1: GetJourNom = "LUN"
        Case 2: GetJourNom = "MAR"
        Case 3: GetJourNom = "MER"
        Case 4: GetJourNom = "JEU"
        Case 5: GetJourNom = "VEN"
        Case 6: GetJourNom = "SAM"
        Case 7: GetJourNom = "DIM"
    End Select
    Exit Function
Err1:
    GetJourNom = ""
End Function

Public Function MatchCouleur(couleurCellule As Long, nomCouleur As String) As Boolean
    MatchCouleur = False
    Dim couleurCible As Long
    
    Select Case UCase(nomCouleur)
        Case "BLEU": couleurCible = 16711680
        Case "BLEU_CLAIR": couleurCible = 16776960
        Case "ROUGE": couleurCible = 255
        Case "JAUNE": couleurCible = 65535
        Case "ORANGE": couleurCible = 49407
        Case "CYAN": couleurCible = 16776960
        Case "ROSE": couleurCible = 13408767
        Case "GRIS": couleurCible = 12632256
        Case Else: Exit Function
    End Select
    
    If Abs(couleurCellule - couleurCible) < 100000 Then MatchCouleur = True
    If couleurCellule = couleurCible Then MatchCouleur = True
End Function
