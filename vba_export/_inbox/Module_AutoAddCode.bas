Attribute VB_Name = "Module_AutoAddCode"
' ExportedAt: 2026-01-21 | Workbook: Planning_2026.xlsm
' MIS A JOUR pour nouvelle structure: F_6h45, F_7h_8h, Matin, PM, Soir, Nuit
Option Explicit

'================================================================================
' STRUCTURE COLONNES Config_Codes (15 colonnes A-O):
'   A = Code
'   B = Description
'   C = Type_Code
'   D = Heures_normales
'   E = TopCode
'   F = H_Start
'   G = H_Pause_Start
'   H = H_Pause_End
'   I = H_End
'   J = F_6h45
'   K = F_7h_8h
'   L = Matin
'   M = PM
'   N = Soir
'   O = Nuit
'================================================================================

'================================================================================
' GERER CODES - Menu principal (Ajouter / Supprimer)
'================================================================================

Public Sub GererCodes()
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Que voulez-vous faire ?" & vbCrLf & vbCrLf & _
                 "OUI = Ajouter un nouveau code" & vbCrLf & _
                 "NON = Supprimer un code existant", _
                 vbQuestion + vbYesNoCancel, "Gestion Codes")
    
    Select Case rep
        Case vbYes
            AjouterNouveauCode
        Case vbNo
            SupprimerCode
        Case vbCancel
            ' Annule
    End Select
End Sub

'================================================================================
' AUTO-ADD CODE - Nouvelle structure 15 colonnes
'================================================================================

Public Sub AjouterNouveauCode()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Feuille Config_Codes introuvable!", vbCritical
        Exit Sub
    End If
    
    ' 1) Demander le code
    Dim code As String
    code = Trim(InputBox("Entrez le code horaire (ex: 8:30 16:30):", "Nouveau Code", ""))
    If code = "" Then Exit Sub
    
    ' 2) Verifier si le code existe deja
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(code) Then
            MsgBox "Ce code existe deja (ligne " & i & ").", vbExclamation
            ws.Activate
            ws.Cells(i, 1).Select
            Exit Sub
        End If
    Next i
    
    ' 3) Parser le code pour obtenir les heures (pour H_Start, H_End, etc.)
    Dim hStart As Double, hEnd As Double
    Dim hPauseStart As Double, hPauseEnd As Double
    Dim codeIsHoraire As Boolean
    codeIsHoraire = ParseHeuresComplet(code, hStart, hPauseStart, hPauseEnd, hEnd)
    
    ' 4) Demander les heures totales
    Dim heuresTotal As Double
    Dim heuresInput As String
    heuresInput = InputBox("Entrez le nombre d'HEURES de travail:" & vbCrLf & _
                           "(ex: 8 pour 8h, 8.5 pour 8h30)", _
                           "Heures Travail", "8")
    If heuresInput = "" Then Exit Sub
    heuresTotal = Val(Replace(heuresInput, ",", "."))
    
    ' 5) Description et type
    Dim description As String, typeCode As String, topCode As String
    description = InputBox("Description du code:", "Description", "Poste de travail")
    If description = "" Then description = "Poste de travail"
    typeCode = "Travail"
    
    ' TopCode?
    Dim repTop As VbMsgBoxResult
    repTop = MsgBox("Ajouter ce code a TopCode (liste deroulante planning)?", vbQuestion + vbYesNo, "TopCode")
    If repTop = vbYes Then
        topCode = "x"
    Else
        topCode = ""
    End If
    
    ' 6) DEMANDER MANUELLEMENT LES FRACTIONS (AVEC SUGGESTIONS INTELLIGENTES)
    MsgBox "Entrez les FRACTIONS." & vbCrLf & _
           "La macro vous propose une valeur par defaut (calculee)." & vbCrLf & _
           "Appuyez sur ENTREE pour accepter, ou tapez votre valeur.", _
           vbInformation, "Fractions Assistees"
    
    ' --- CALCUL DES SUGGESTIONS (V3 - INTELLIGENCE ETENDUE) ---
    Dim defF6h45 As String, defF7h8h As String
    Dim defMatin As String, defPM As String, defSoir As String, defNuit As String
    
    ' F_6h45: Selon CSV, Start <= 6:45 donne 1
    defF6h45 = IIf(hStart <= 6.75, "1", "")
    
    ' F_7h_8h: Selon CSV, Start < 8 et End > 7
    defF7h8h = IIf(hStart < 8 And hEnd > 7, "1", "")
    
    ' MATIN: Selon CSV, Start < 12
    defMatin = IIf(hStart < 12, "1", "")
    
    ' PM: INTELLIGENCE 'A LA CARTE'
    ' 1. Les Codes Coupés 'C ...' n'ont pas de PM (pause midi)
    If Left(UCase(code), 1) = "C" Then
        defPM = "" 
    ' 2. Les PM Courts (ex: 8-14h) -> 0,5
    ElseIf hEnd > 12 And hEnd <= 14.5 Then
        defPM = "0,5"
    ' 3. Les PM Classiques -> 1
    ElseIf hEnd > 12 Then
        defPM = "1"
    Else
        defPM = ""
    End If
    
    ' SOIR: REGLES DEDUITES + DEMANDE UTILISATEUR
    ' CSV: 15h30 n'a pas de Soir, 16h a un Soir.
    ' Demande: Jusqu'à 17h30 -> 0,5
    If hEnd <= 15.5 Then
        defSoir = ""
    ElseIf hEnd <= 17.5 Then
        defSoir = "0,5" ' ex: 16h, 16h30, 17h, 17h30 -> demi-Soir
    Else
        defSoir = "1" ' > 17h30 -> Soir plein
    End If
    
    ' NUIT: REGLES DEDUITES
    ' Start >= 19:45 (ex: 19h45, 20h) OU End <= 8 (nuit complète)
    If hStart >= 19.75 Or hEnd <= 8 Then
        ' Exception: "20 24" -> Demi-nuit (CSV ligne 85)
        ' Si ça finit à 24h (0h) -> 0,5
        If hEnd = 0 Or hEnd = 24 Then
            defNuit = "0,5"
        Else
            defNuit = "1"
        End If
    Else
        defNuit = ""
    End If
    
    
    ' --- SAISIE UTILISATEUR (avec defaut) ---
    Dim f6h45 As String, f7h8h As String
    Dim pMatin As String, pPM As String, pSoir As String, pNuit As String
    
    f6h45 = InputBox("F_6h45 (present a 6h45):", "F_6h45", defF6h45)
    f7h8h = InputBox("F_7h_8h (present entre 7h et 8h):", "F_7h_8h", defF7h8h)
    pMatin = InputBox("MATIN (travaille le matin):", "Matin", defMatin)
    pPM = InputBox("PM (travaille l'apres-midi):" & vbCrLf & "Suggestion: 0,5 si <= 14h30, Vide si Coupe", "PM", defPM)
    pSoir = InputBox("SOIR (finit apres 16h30):" & vbCrLf & "Suggestion: 0,5 si <= 17h30, 1 si > 17h30", "Soir", defSoir)
    pNuit = InputBox("NUIT (poste de nuit):" & vbCrLf & "Suggestion: 0,5 si demi-nuit (20-24h)", "Nuit", defNuit)
    
    ' 7) Nouvelle ligne
    Dim newRow As Long
    newRow = lastRow + 1
    
    ' 8) Remplir la ligne (15 colonnes A-O)
    Application.ScreenUpdating = False
    
    ws.Cells(newRow, 1).Value = code                           ' A: Code
    ws.Cells(newRow, 2).Value = description                    ' B: Description
    ws.Cells(newRow, 3).Value = typeCode                       ' C: Type_Code
    ws.Cells(newRow, 4).Value = heuresTotal                    ' D: Heures_normales
    ws.Cells(newRow, 5).Value = topCode                        ' E: TopCode
    ws.Cells(newRow, 6).Value = FormaterHeure(hStart)          ' F: H_Start
    ws.Cells(newRow, 7).Value = FormaterHeure(hPauseStart)     ' G: H_Pause_Start
    ws.Cells(newRow, 8).Value = FormaterHeure(hPauseEnd)       ' H: H_Pause_End
    ws.Cells(newRow, 9).Value = FormaterHeure(hEnd)            ' I: H_End
    ws.Cells(newRow, 10).Value = f6h45                         ' J: F_6h45
    ws.Cells(newRow, 11).Value = f7h8h                         ' K: F_7h_8h
    ws.Cells(newRow, 12).Value = pMatin                        ' L: Matin
    ws.Cells(newRow, 13).Value = pPM                           ' M: PM
    ws.Cells(newRow, 14).Value = pSoir                         ' N: Soir
    ws.Cells(newRow, 15).Value = pNuit                         ' O: Nuit
    
    Application.ScreenUpdating = True
    
    ' 9) Afficher resume
    MsgBox "Code '" & code & "' ajoute!" & vbCrLf & vbCrLf & _
           "Heures: " & heuresTotal & "h" & vbCrLf & _
           "F_6h45: " & f6h45 & " | F_7h_8h: " & f7h8h & vbCrLf & _
           "Matin: " & pMatin & " | PM: " & pPM & " | Soir: " & pSoir & " | Nuit: " & pNuit, _
           vbInformation, "Code Ajoute"
    
    ' 10) Trier et regenerer liste
    TrierCodesParHeure
    GenererListeCodesDropdown
End Sub

'================================================================================
' SUPPRIMER UN CODE
'================================================================================

Public Sub SupprimerCode()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Feuille Config_Codes introuvable!", vbCritical
        Exit Sub
    End If
    
    Dim code As String
    code = Trim(InputBox("Entrez le code a supprimer:", "Supprimer Code", ""))
    If code = "" Then Exit Sub
    
    Dim lastRow As Long, i As Long
    Dim found As Boolean, foundRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    found = False
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(code) Then
            found = True
            foundRow = i
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "Code '" & code & "' non trouve.", vbExclamation
        Exit Sub
    End If
    
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Supprimer le code '" & code & "' (ligne " & foundRow & ") ?", _
                 vbQuestion + vbYesNo, "Confirmer Suppression")
    
    If rep = vbYes Then
        ws.Rows(foundRow).Delete
        MsgBox "Code '" & code & "' supprime.", vbInformation
        GenererListeCodesDropdown
    End If
End Sub

'================================================================================
' TRIER PAR HEURE (colonne F = H_Start, puis I = H_End)
'================================================================================

Public Sub TrierCodesParHeure()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Sub
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 9 Then lastCol = 15
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    ' Trier par H_Start (colonne F) puis H_End (colonne I)
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Sort _
        Key1:=ws.Range("F2"), Order1:=xlAscending, _
        Key2:=ws.Range("I2"), Order2:=xlAscending, _
        Header:=xlNo
    On Error GoTo 0
    
    Application.ScreenUpdating = True
End Sub

'================================================================================
' PARSE HEURES COMPLET (gere horaires coupes)
'================================================================================

Private Function ParseHeuresComplet(ByVal code As String, ByRef hStart As Double, _
    ByRef hPauseStart As Double, ByRef hPauseEnd As Double, ByRef hEnd As Double) As Boolean
    
    On Error GoTo Echec
    
    code = Replace(code, "-", " ")
    code = Trim(code)
    
    Dim parts() As String
    Dim numParts() As String
    Dim p As Variant
    Dim cnt As Long
    
    parts = Split(code, " ")
    ReDim numParts(0 To UBound(parts))
    cnt = 0
    
    For Each p In parts
        If Len(CStr(p)) > 0 Then
            If IsNumeric(Left(CStr(p), 1)) Then
                numParts(cnt) = CStr(p)
                cnt = cnt + 1
            End If
        End If
    Next p
    
    If cnt < 2 Then GoTo Echec
    
    hStart = ConvertirHeure(numParts(0))
    hPauseStart = 0
    hPauseEnd = 0
    
    If cnt = 2 Then
        hEnd = ConvertirHeure(numParts(1))
    ElseIf cnt = 4 Then
        hPauseStart = ConvertirHeure(numParts(1))
        hPauseEnd = ConvertirHeure(numParts(2))
        hEnd = ConvertirHeure(numParts(3))
    Else
        hEnd = ConvertirHeure(numParts(cnt - 1))
    End If
    
    ' Gerer nuit (fin avant debut)
    If hEnd <= hStart And hEnd < 12 Then hEnd = hEnd + 24
    
    ParseHeuresComplet = True
    Exit Function
    
Echec:
    hStart = 0: hPauseStart = 0: hPauseEnd = 0: hEnd = 0
    ParseHeuresComplet = False
End Function

'================================================================================
' CALCULER HEURES TOTAL
'================================================================================

Private Function CalculerHeuresTotal(ByVal code As String) As Double
    Dim total As Double
    Dim parts() As String
    Dim numParts() As String
    Dim p As Variant
    Dim cnt As Long, i As Long
    Dim h1 As Double, h2 As Double
    
    code = Replace(code, "-", " ")
    code = Trim(code)
    parts = Split(code, " ")
    ReDim numParts(0 To UBound(parts))
    cnt = 0
    
    For Each p In parts
        If Len(CStr(p)) > 0 Then
            If IsNumeric(Left(CStr(p), 1)) Then
                numParts(cnt) = CStr(p)
                cnt = cnt + 1
            End If
        End If
    Next p
    
    If cnt < 2 Or cnt Mod 2 <> 0 Then
        CalculerHeuresTotal = 0
        Exit Function
    End If
    
    total = 0
    For i = 0 To cnt - 1 Step 2
        h1 = ConvertirHeure(numParts(i))
        h2 = ConvertirHeure(numParts(i + 1))
        If h2 <= h1 Then h2 = h2 + 24
        total = total + (h2 - h1)
    Next i
    
    CalculerHeuresTotal = total
End Function

'================================================================================
' CONVERTIR HEURE
'================================================================================

Private Function ConvertirHeure(ByVal s As String) As Double
    s = Trim(s)
    Dim i As Long, c As String, clean As String
    
    clean = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If IsNumeric(c) Or c = ":" Or c = "." Or c = "," Then
            clean = clean & c
        Else
            Exit For
        End If
    Next i
    
    clean = Replace(clean, ",", ".")
    
    If InStr(clean, ":") > 0 Then
        Dim p() As String
        p = Split(clean, ":")
        ConvertirHeure = Val(p(0)) + Val(p(1)) / 60
    Else
        ConvertirHeure = Val(clean)
    End If
End Function

'================================================================================
' FORMATER HEURE (avec secondes pour compatibilite tri)
'================================================================================

Private Function FormaterHeure(h As Double) As String
    If h = 0 Then
        FormaterHeure = ""
        Exit Function
    End If
    
    ' Gerer heures > 24 (nuit)
    If h >= 24 Then h = h - 24
    
    Dim hrs As Long, mins As Long
    hrs = Int(h)
    mins = Int((h - hrs) * 60 + 0.5)
    
    If mins = 60 Then
        hrs = hrs + 1
        mins = 0
    End If
    
    ' Format avec secondes pour compatibilite tri Excel
    FormaterHeure = Format(hrs, "00") & ":" & Format(mins, "00") & ":00"
End Function

'================================================================================
' RECALCULER TOUS LES CODES - Met a jour F_6h45, F_7h_8h, Matin, PM, Soir, Nuit
'================================================================================

Public Sub RecalculerTousLesCodes_DESACTIVE()
    MsgBox "Cette macro a ete desactivee pour preserve les modifications manuelles des fractions.", vbInformation
    Exit Sub
End Sub

'================================================================================
' OUTIL DE CORRECTION : Enlever les PM des codes Coupes existants
'================================================================================
Public Sub Corriger_PM_Coupes_Existants()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config_Codes")
    
    Dim lastRow As Long, i As Long
    Dim code As String, count As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRow
        code = UCase(Trim(ws.Cells(i, 1).Value))
        ' Si c'est un code qui commence par C (C 15, C 19, C 20...)
        ' ET qu'il a un PM coche (col 13 / M)
        If Left(code, 1) = "C" And ws.Cells(i, 13).Value = "1" Then
            ws.Cells(i, 13).Value = "" ' On efface le PM
            count = count + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Correction terminee !" & vbCrLf & _
           count & " codes 'Coupe' ont vu leur PM retire.", vbInformation
End Sub
