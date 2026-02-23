Attribute VB_Name = "Module_HeuresTravaillees"
'====================================================================
' MODULE HEURES TRAVAILLEES - Planning_2026_RUNTIME
' Calcul des heures theoriques vs prestees par agent/mois
' Gestion: temps partiels, nuits, feries belges, codes fractionnaires
'
' DEPENDANCES:
'   - Module_Planning_Core: BuildFeriesBE(), ParseCode(), HeureDec()
'   - Feuille "Personnel": HeuresStdJour (C4), %Temps mensuels
'   - Feuilles mensuelles: Janv..Dec (Row 6+ = agents, Col 3-33 = jours)
'
' INSTALLATION:
'   1. Ouvrir Planning_2026_RUNTIME.xlsm
'   2. Alt+F11 > Menu Fichier > Importer > Module_HeuresTravaillees.bas
'   3. Verifier que Module_Planning_Core est present
'   4. Lancer: Alt+F8 > RecalculerHeuresMois > Executer
'====================================================================

Option Explicit

' ---- CONSTANTES ----
Private Const MONTH_SHEETS As String = "Janv,Fev,Mars,Avril,Mai,Juin,Juil,Aout,Sept,Oct,Nov,Dec"
Private Const FIRST_EMP_ROW As Long = 5     ' Premiere ligne agent
Private Const FIRST_DAY_COL As Long = 3     ' Colonne C = jour 1
Private Const MAX_DAY_COL As Long = 33      ' Colonne AG = jour 31

' Colonnes Personnel pour % mensuel (C29=Janv%, C31=Fev%, etc. = +2 par mois)
Private Const PERSONNEL_PCT_START_COL As Long = 29  ' Colonne pour Janvier %
Private Const PERSONNEL_PCT_STEP As Long = 2        ' +2 colonnes par mois

' ---- HELPER: conversion safe cellule -> String (evite Type Mismatch sur erreurs #N/A etc.) ----
Private Function SafeCStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeCStr = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeCStr = ""
    Else
        SafeCStr = CStr(v)
    End If
End Function

' ---- CODES D'ABSENCE RECONNUS (ne comptent PAS comme heures prestees) ----
' Ces codes sont traites separement ou ignores dans le calcul des heures effectives
Private Const ABSENCE_CODES As String = _
    "CA,EL,ANC,C SOC,DP,CTR,RCT,RV,WE,DECES,RHS,JF"

' ---- CODES DE MALADIE ----
Private Const MALADIE_PREFIXES As String = "MAL-,MUT,MAT-,PAT-"

'====================================================================
' FONCTION PUBLIQUE 1: JoursOuvrablesMois
' Compte les jours ouvrables (Lun-Ven) moins les jours feries belges
'
' @param annee   Long   Annee (ex: 2026)
' @param moisNum Long   Numero du mois (1-12)
' @return Long   Nombre de jours ouvrables
'====================================================================
Public Function JoursOuvrablesMois(ByVal annee As Long, ByVal moisNum As Long) As Long
    Dim compteur As Long
    Dim premierJour As Date
    Dim dernierJour As Date
    Dim d As Date
    Dim feries As Object ' Dictionary

    ' Validation
    If moisNum < 1 Or moisNum > 12 Then
        JoursOuvrablesMois = 0
        Exit Function
    End If

    ' Bornes du mois
    premierJour = DateSerial(annee, moisNum, 1)
    dernierJour = DateSerial(annee, moisNum + 1, 0) ' Dernier jour du mois

    ' Recuperer les jours feries belges via Module_Planning_Core
    Set feries = Module_Planning_Core.BuildFeriesBE(annee)

    compteur = 0
    For d = premierJour To dernierJour
        ' Exclure samedi (7) et dimanche (1) - Weekday avec vbMonday = Lun=1..Dim=7
        If Weekday(d, vbMonday) <= 5 Then
            ' Exclure les jours feries
            If Not feries.Exists(CStr(d)) Then
                compteur = compteur + 1
            End If
        End If
    Next d

    JoursOuvrablesMois = compteur
End Function

'====================================================================
' FONCTION PUBLIQUE 2: HeuresTheoriquesMois
' Heures theoriques = jours ouvrables * heuresStdJour * pctTemps
'
' @param annee         Long     Annee (ex: 2026)
' @param moisNum       Long     Numero du mois (1-12)
' @param pctTemps      Double   Pourcentage temps de travail (0.75, 1.0, etc.)
' @param heuresStdJour Double   Heures standard par jour (7.6, 5.7, 6.08, etc.)
' @return Double   Heures theoriques pour le mois
'====================================================================
Public Function HeuresTheoriquesMois(ByVal annee As Long, _
                                      ByVal moisNum As Long, _
                                      ByVal pctTemps As Double, _
                                      ByVal heuresStdJour As Double) As Double
    Dim joursOuvrables As Long

    ' Validation
    If pctTemps <= 0 Or pctTemps > 1.5 Then
        ' Securite: si le % semble aberrant, on force 100%
        pctTemps = 1#
    End If
    If heuresStdJour <= 0 Or heuresStdJour > 12 Then
        ' Securite: valeur par defaut secteur hospitalier belge
        heuresStdJour = 7.6
    End If

    joursOuvrables = JoursOuvrablesMois(annee, moisNum)

    ' Formule: jours ouvrables * heures standard/jour
    ' Le pctTemps est deja reflete dans heuresStdJour pour les temps partiels
    ' (un 75% a un heuresStdJour de 5.7 = 7.6 * 0.75)
    ' Mais si les deux sont fournis separement, on multiplie
    HeuresTheoriquesMois = joursOuvrables * heuresStdJour
End Function

'====================================================================
' FONCTION PUBLIQUE 3: DureeEffectiveCode
' Convertit un code horaire en nombre d'heures decimales
'
' Gere:
'   - Plages horaires: "7 15:30" -> 8.5h, "8:30 16:30" -> 8h
'   - Nuits (passage minuit): "19:45 6:45" -> 11h
'   - Codes coupes: "C 15" -> calcul special, "C 19" -> idem
'   - Fractions: "3/4*" -> 75% du std, "4/5*" -> 80% du std
'   - Absences: "CA", "EL", "WE" etc. -> 0h (pas des heures prestees)
'
' @param code           String   Le code de la cellule planning
' @param heuresStdJour  Double   (Optionnel) Heures std/jour pour fractions (defaut 7.6)
' @return Double   Duree en heures decimales
'====================================================================
Public Function DureeEffectiveCode(ByVal code As String, _
                                    Optional ByVal heuresStdJour As Double = 7.6) As Double
    Dim c As String
    Dim result As Double

    c = Trim(code)
    If Len(c) = 0 Or c = "0" Then
        DureeEffectiveCode = 0#
        Exit Function
    End If

    ' ---- 1. Codes d'absence -> 0 heures prestees ----
    If EstCodeAbsence(c) Then
        DureeEffectiveCode = 0#
        Exit Function
    End If

    ' ---- 2. Codes de maladie -> 0 heures prestees ----
    If EstCodeMaladie(c) Then
        DureeEffectiveCode = 0#
        Exit Function
    End If

    ' ---- 3. Codes feries (F-xxx, R-xxx) -> 0 heures prestees ----
    If EstCodeFerie(c) Then
        DureeEffectiveCode = 0#
        Exit Function
    End If

    ' ---- 4. Codes fractionnaires: "3/4*", "4/5*", "1/2*" ----
    result = ParseFraction(c, heuresStdJour)
    If result > 0 Then
        DureeEffectiveCode = result
        Exit Function
    End If

    ' ---- 5. Codes coupes: "C 15", "C 19", "C 20", "C 20 E" ----
    result = ParseCodeCoupe(c)
    If result > 0 Then
        DureeEffectiveCode = result
        Exit Function
    End If

    ' ---- 6. Plages horaires: "7 15:30", "8:30 16:30", "19:45 6:45" ----
    result = ParsePlageHoraire(c)
    If result > 0 Then
        DureeEffectiveCode = result
        Exit Function
    End If

    ' ---- 7. Code non reconnu -> 0 ----
    DureeEffectiveCode = 0#
End Function

'====================================================================
' FONCTION PUBLIQUE 4: HeuresPresteesMois
' Parse chaque cellule d'une ligne agent dans un onglet mois
' et somme les heures effectives
'
' @param nomMois   String   Nom de l'onglet mois ("Janv", "Fev", etc.)
' @param rowIndex  Long     Numero de ligne de l'agent dans la feuille
' @return Double   Total heures prestees sur le mois
'====================================================================
Public Function HeuresPresteesMois(ByVal nomMois As String, _
                                    ByVal rowIndex As Long) As Double
    Dim ws As Worksheet
    Dim totalHeures As Double
    Dim col As Long
    Dim cellVal As String
    Dim numJours As Long
    Dim heuresStd As Double

    ' Recuperer la feuille
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then
        HeuresPresteesMois = 0#
        Exit Function
    End If

    ' Recuperer le heuresStdJour de l'agent depuis la feuille Personnel
    heuresStd = GetHeuresStdJourAgent(ws.Cells(rowIndex, 1).Value)
    If heuresStd <= 0 Then heuresStd = 7.6 ' Fallback

    ' Determiner le nombre de jours dans le mois (depuis les en-tetes)
    numJours = CompterJoursMois(ws)

    totalHeures = 0#
    For col = FIRST_DAY_COL To FIRST_DAY_COL + numJours - 1
        cellVal = Trim(SafeCStr(ws.Cells(rowIndex, col).Value))
        If Len(cellVal) > 0 And cellVal <> "0" Then
            totalHeures = totalHeures + DureeEffectiveCode(cellVal, heuresStd)
        End If
    Next col

    HeuresPresteesMois = Round(totalHeures, 2)
End Function

'====================================================================
' SUB PUBLIQUE 5: RecalculerHeuresMois
' Pour tous les agents d'un mois donne, calcule:
'   - Heures theoriques (basees sur jours ouvrables et contrat)
'   - Heures prestees (somme des plages horaires)
'   - Delta (prestees - theoriques)
'
' Ecrit les resultats dans des colonnes dediees a droite du planning
'
' @param nomMois   String   Nom de l'onglet ("Janv", "Fev", etc.)
'                           Si vide, traite TOUS les mois
'====================================================================
Public Sub RecalculerHeuresMois(Optional ByVal nomMois As String = "")
    If Len(nomMois) > 0 Then
        RecalculerUnMois nomMois
    Else
        ' Traiter tous les mois
        Dim mSheets() As String
        Dim m As Long
        mSheets = Split(MONTH_SHEETS, ",")

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        For m = 0 To 11
            RecalculerUnMois mSheets(m)
        Next m

        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
End Sub

'====================================================================
' SOUS-ROUTINE: RecalculerUnMois (Private)
' Traite un seul mois
'====================================================================
Private Sub RecalculerUnMois(ByVal nomMois As String)
    Dim ws As Worksheet
    Dim annee As Long
    Dim moisNum As Long
    Dim numJours As Long
    Dim colTheo As Long, colPrest As Long, colDelta As Long
    Dim r As Long, lastR As Long
    Dim agName As String
    Dim heuresStd As Double
    Dim pctTemps As Double
    Dim hTheo As Double, hPrest As Double, delta As Double

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    annee = GetAnneeFromSheet(ws)
    If annee = 0 Then annee = 2026
    moisNum = GetMoisNumFromName(nomMois)
    If moisNum = 0 Then Exit Sub
    numJours = CompterJoursMois(ws)

    colTheo = FIRST_DAY_COL + numJours + 1
    colPrest = colTheo + 1
    colDelta = colPrest + 1

    ' Desactiver events pour eviter cascade Worksheet_Change
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error GoTo CleanupEvents

    ' En-tetes
    ws.Cells(5, colTheo).Value = "H.Theo"
    ws.Cells(5, colTheo).Font.Bold = True
    ws.Cells(5, colTheo).Interior.Color = RGB(198, 224, 180)
    ws.Cells(5, colPrest).Value = "H.Prest"
    ws.Cells(5, colPrest).Font.Bold = True
    ws.Cells(5, colPrest).Interior.Color = RGB(180, 198, 231)
    ws.Cells(5, colDelta).Value = "Delta"
    ws.Cells(5, colDelta).Font.Bold = True
    ws.Cells(5, colDelta).Interior.Color = RGB(255, 230, 153)

    lastR = GetLastEmployeeRow(ws)
    For r = FIRST_EMP_ROW To lastR
        If IsError(ws.Cells(r, 1).Value) Then GoTo NextAgent
        agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
        If Len(agName) = 0 Then GoTo NextAgent
        If InStr(agName, "Remplacement") > 0 Then GoTo NextAgent
        If agName = "Us Nuit" Then GoTo NextAgent

        heuresStd = GetHeuresStdJourAgent(agName)
        If heuresStd <= 0 Then heuresStd = 7.6

        pctTemps = GetPctTempsAgent(agName, moisNum)
        If pctTemps <= 0 Then pctTemps = 1#

        hTheo = HeuresTheoriquesMois(annee, moisNum, pctTemps, heuresStd)
        hPrest = HeuresPresteesMois(nomMois, r)

        ws.Cells(r, colTheo).Value = Round(hTheo, 2)
        ws.Cells(r, colTheo).NumberFormat = "0.00"
        ws.Cells(r, colPrest).Value = Round(hPrest, 2)
        ws.Cells(r, colPrest).NumberFormat = "0.00"

        delta = Round(hPrest - hTheo, 2)
        ws.Cells(r, colDelta).Value = delta
        ws.Cells(r, colDelta).NumberFormat = "0.00"

        If delta < -2 Then
            ws.Cells(r, colDelta).Font.Color = RGB(204, 0, 0)
            ws.Cells(r, colDelta).Font.Bold = True
        ElseIf delta > 2 Then
            ws.Cells(r, colDelta).Font.Color = RGB(0, 128, 0)
            ws.Cells(r, colDelta).Font.Bold = True
        Else
            ws.Cells(r, colDelta).Font.Color = RGB(0, 0, 0)
            ws.Cells(r, colDelta).Font.Bold = False
        End If
NextAgent:
    Next r

    ws.Columns(colTheo).ColumnWidth = 8
    ws.Columns(colPrest).ColumnWidth = 8
    ws.Columns(colDelta).ColumnWidth = 8

CleanupEvents:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'====================================================================
' FONCTIONS PRIVEES DE PARSING
'====================================================================

'--------------------------------------------------------------------
' ParsePlageHoraire: "7 15:30" -> 8.5, "19:45 6:45" -> 11.0
' Detecte automatiquement les nuits (debut > fin = passage minuit)
'--------------------------------------------------------------------
Private Function ParsePlageHoraire(ByVal code As String) As Double
    Dim parts() As String
    Dim hDebut As Double, hFin As Double
    Dim duree As Double
    Dim c As String

    c = Trim(code)

    ' Le code doit contenir au moins un espace separant debut et fin
    ' Formats possibles: "7 15:30", "8:30 16:30", "19:45 6:45", "7 15"
    ' On cherche le pattern: [hh[:mm]] [hh[:mm]]

    ' Nettoyer les espaces multiples
    Do While InStr(c, "  ") > 0
        c = Replace(c, "  ", " ")
    Loop

    parts = Split(c, " ")

    ' On a besoin d'exactement 2 parties (debut fin)
    ' Mais certains codes comme "C 15" ont aussi 2 parties -> verifier
    If UBound(parts) < 1 Then
        ParsePlageHoraire = 0#
        Exit Function
    End If

    ' Essayer de parser les deux parties comme des heures
    hDebut = ParseHeure(parts(0))
    hFin = ParseHeure(parts(UBound(parts)))

    ' Si l'une des deux n'est pas parsable -> pas une plage horaire
    If hDebut < 0 Or hFin < 0 Then
        ParsePlageHoraire = 0#
        Exit Function
    End If

    ' Calcul de la duree
    If hFin > hDebut Then
        ' Shift normal (ex: 7h -> 15h30)
        duree = hFin - hDebut
    ElseIf hFin < hDebut Then
        ' Shift de nuit passant minuit (ex: 19h45 -> 6h45)
        duree = (24# - hDebut) + hFin
    Else
        ' Debut = Fin -> 24h ou erreur, on met 0
        duree = 0#
    End If

    ' Securite: une prestation de plus de 16h est probablement une erreur
    If duree > 16 Then duree = 0#

    ParsePlageHoraire = Round(duree, 2)
End Function

'--------------------------------------------------------------------
' ParseHeure: "15:30" -> 15.5, "7" -> 7.0, "19:45" -> 19.75
' Retourne -1 si non parsable
'--------------------------------------------------------------------
Private Function ParseHeure(ByVal txt As String) As Double
    Dim parts() As String
    Dim h As Double, m As Double

    txt = Trim(txt)
    If Len(txt) = 0 Then
        ParseHeure = -1#
        Exit Function
    End If

    If InStr(txt, ":") > 0 Then
        ' Format HH:MM
        parts = Split(txt, ":")
        If UBound(parts) <> 1 Then
            ParseHeure = -1#
            Exit Function
        End If
        If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then
            ParseHeure = -1#
            Exit Function
        End If
        h = CDbl(parts(0))
        m = CDbl(parts(1))
        If h < 0 Or h > 23 Or m < 0 Or m > 59 Then
            ParseHeure = -1#
            Exit Function
        End If
        ParseHeure = h + (m / 60#)
    Else
        ' Format H ou HH (entier seul)
        If Not IsNumeric(txt) Then
            ParseHeure = -1#
            Exit Function
        End If
        h = CDbl(txt)
        If h < 0 Or h > 23 Then
            ParseHeure = -1#
            Exit Function
        End If
        ParseHeure = h
    End If
End Function

'--------------------------------------------------------------------
' ParseFraction: "3/4*" -> 75% de heuresStdJour, "4/5*" -> 80%
' Retourne 0 si ce n'est pas un code fraction
'--------------------------------------------------------------------
Private Function ParseFraction(ByVal code As String, ByVal heuresStdJour As Double) As Double
    Dim c As String
    Dim parts() As String
    Dim numerateur As Double, denominateur As Double

    c = Trim(code)
    ParseFraction = 0#

    ' Les codes fraction finissent par "*" et contiennent "/"
    If Right(c, 1) <> "*" Then Exit Function
    If InStr(c, "/") = 0 Then Exit Function

    ' Retirer le "*" final
    c = Left(c, Len(c) - 1)

    parts = Split(c, "/")
    If UBound(parts) <> 1 Then Exit Function

    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Function

    numerateur = CDbl(parts(0))
    denominateur = CDbl(parts(1))

    If denominateur = 0 Then Exit Function

    ' Fraction * heures standard du jour
    ParseFraction = Round((numerateur / denominateur) * heuresStdJour, 2)
End Function

'--------------------------------------------------------------------
' ParseCodeCoupe: "C 15" -> prestation coupee
' Convention hospitaliere belge:
'   "C 15" = Coupe finissant a 15h (typiquement 7h-11h + 12h-15h = 7h)
'   "C 19" = Coupe finissant a 19h (typiquement 7h-13h + 15h-19h = 10h)
'   "C 20" = Coupe finissant a 20h (typiquement 7h-13h + 15h-20h = 11h)
'   "C 20 E" = Coupe etendu finissant a 20h
'
' Logique: on estime la duree selon l'heure de fin
' avec une pause standard d'1h deduite
'--------------------------------------------------------------------
Private Function ParseCodeCoupe(ByVal code As String) As Double
    Dim c As String
    Dim parts() As String
    Dim hFin As Double
    Dim hDebut As Double
    Dim pauseStd As Double

    c = Trim(UCase(code))
    ParseCodeCoupe = 0#

    ' Doit commencer par "C "
    If Left(c, 2) <> "C " Then Exit Function

    ' Nettoyer
    c = Mid(c, 3) ' Retirer "C "
    c = Trim(c)

    ' Retirer suffixes ("E" pour etendu, etc.)
    c = Replace(c, " E", "")
    c = Replace(c, "E", "")
    c = Trim(c)

    ' Parser l'heure de fin
    If Not IsNumeric(c) Then
        ' Peut contenir ":" -> essayer
        hFin = ParseHeure(c)
        If hFin < 0 Then Exit Function
    Else
        hFin = CDbl(c)
    End If

    ' Heure de debut standard pour les coupes: 7h
    hDebut = 7#

    ' Pause standard: 1h (conventions secteur hospitalier belge)
    pauseStd = 1#

    ' Duree = (fin - debut) - pause
    If hFin > hDebut Then
        ParseCodeCoupe = hFin - hDebut - pauseStd
    Else
        ' Cas improbable: passage minuit pour un coupe
        ParseCodeCoupe = (24# - hDebut) + hFin - pauseStd
    End If

    ' Securite
    If ParseCodeCoupe < 0 Then ParseCodeCoupe = 0#
    If ParseCodeCoupe > 16 Then ParseCodeCoupe = 0#
End Function

'====================================================================
' FONCTIONS PRIVEES DE CLASSIFICATION
'====================================================================

'--------------------------------------------------------------------
' EstCodeAbsence: verifie si le code est un code d'absence connu
'--------------------------------------------------------------------
Private Function EstCodeAbsence(ByVal code As String) As Boolean
    Dim c As String
    Dim absCodes() As String
    Dim i As Long

    c = Trim(UCase(code))
    absCodes = Split(ABSENCE_CODES, ",")

    For i = 0 To UBound(absCodes)
        If c = Trim(absCodes(i)) Then
            EstCodeAbsence = True
            Exit Function
        End If
    Next i

    EstCodeAbsence = False
End Function

'--------------------------------------------------------------------
' EstCodeMaladie: verifie les prefixes maladie
' MAL-GAR, MAL-MUT, MUT, MAT-EMP, MAT-MUT, PAT-EMP, PAT-MUT
'--------------------------------------------------------------------
Private Function EstCodeMaladie(ByVal code As String) As Boolean
    Dim c As String
    Dim prefixes() As String
    Dim i As Long

    c = Trim(UCase(code))
    prefixes = Split(MALADIE_PREFIXES, ",")

    For i = 0 To UBound(prefixes)
        If Left(c, Len(Trim(prefixes(i)))) = Trim(prefixes(i)) Then
            EstCodeMaladie = True
            Exit Function
        End If
    Next i

    EstCodeMaladie = False
End Function

'--------------------------------------------------------------------
' EstCodeFerie: "F-xxx" ou "R-xxx" = code ferie/recup ferie
'--------------------------------------------------------------------
Private Function EstCodeFerie(ByVal code As String) As Boolean
    Dim c As String
    c = Trim(UCase(code))

    If Len(c) >= 2 Then
        If (Left(c, 1) = "F" Or Left(c, 1) = "R") And Mid(c, 2, 1) = "-" Then
            EstCodeFerie = True
            Exit Function
        End If
    End If

    EstCodeFerie = False
End Function

'====================================================================
' FONCTIONS PRIVEES DE LECTURE FEUILLE
'====================================================================

'--------------------------------------------------------------------
' CompterJoursMois: compte le nombre de jours dans un mois
' en se basant sur les en-tetes de ligne 4 (numeros de jours)
'--------------------------------------------------------------------
Private Function CompterJoursMois(ByVal ws As Worksheet) As Long
    Dim col As Long
    Dim compteur As Long

    compteur = 0
    For col = FIRST_DAY_COL To MAX_DAY_COL
        If IsNumeric(ws.Cells(4, col).Value) And _
           Len(Trim(SafeCStr(ws.Cells(4, col).Value))) > 0 Then
            compteur = compteur + 1
        Else
            ' Premiere cellule vide apres les numeros = fin du mois
            If compteur > 0 Then Exit For
        End If
    Next col

    ' Securite: un mois a entre 28 et 31 jours
    If compteur < 28 Then compteur = 31
    If compteur > 31 Then compteur = 31

    CompterJoursMois = compteur
End Function

'--------------------------------------------------------------------
' GetLastEmployeeRow: trouve la derniere ligne agent
' (cherche en descendant jusqu'a trouver une ligne vide)
'--------------------------------------------------------------------
Private Function GetLastEmployeeRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim emptyCount As Long
    Dim lastFound As Long
    lastFound = FIRST_EMP_ROW - 1
    emptyCount = 0
    For r = FIRST_EMP_ROW To 200
        If Len(Trim(SafeCStr(ws.Cells(r, 1).Value))) > 0 Then
            lastFound = r
            emptyCount = 0
        Else
            emptyCount = emptyCount + 1
            If emptyCount >= 5 Then Exit For
        End If
    Next r
    GetLastEmployeeRow = lastFound
End Function

'--------------------------------------------------------------------
' GetAnneeFromSheet: essaie de recuperer l'annee depuis la feuille
' Cherche dans la cellule A1 ou A2 un texte contenant une annee
'--------------------------------------------------------------------
Private Function GetAnneeFromSheet(ByVal ws As Worksheet) As Long
    Dim val1 As String, val2 As String

    val1 = SafeCStr(ws.Cells(1, 1).Value)
    val2 = SafeCStr(ws.Cells(2, 1).Value)

    ' Chercher un nombre a 4 chiffres commencant par 20
    GetAnneeFromSheet = ExtractYear(val1)
    If GetAnneeFromSheet > 0 Then Exit Function

    GetAnneeFromSheet = ExtractYear(val2)
    If GetAnneeFromSheet > 0 Then Exit Function

    ' Fallback: annee en cours
    GetAnneeFromSheet = Year(Now)
End Function

Private Function ExtractYear(ByVal txt As String) As Long
    Dim i As Long
    Dim chunk As String

    ExtractYear = 0
    For i = 1 To Len(txt) - 3
        chunk = Mid(txt, i, 4)
        If IsNumeric(chunk) Then
            If CLng(chunk) >= 2020 And CLng(chunk) <= 2035 Then
                ExtractYear = CLng(chunk)
                Exit Function
            End If
        End If
    Next i
End Function

'--------------------------------------------------------------------
' GetMoisNumFromName: convertit nom d'onglet en numero de mois
'--------------------------------------------------------------------
Private Function GetMoisNumFromName(ByVal nomMois As String) As Long
    Dim mSheets() As String
    Dim i As Long

    mSheets = Split(MONTH_SHEETS, ",")
    For i = 0 To 11
        If UCase(Trim(mSheets(i))) = UCase(Trim(nomMois)) Then
            GetMoisNumFromName = i + 1
            Exit Function
        End If
    Next i

    GetMoisNumFromName = 0
End Function

'====================================================================
' FONCTIONS PRIVEES DE LECTURE PERSONNEL
'====================================================================

'--------------------------------------------------------------------
' GetHeuresStdJourAgent: recupere heuresStdJour depuis feuille Personnel
' Cherche le nom dans la colonne A, retourne la valeur en C4
' (ou la valeur specifique a l'agent si elle est sur sa ligne)
'--------------------------------------------------------------------
Private Function GetHeuresStdJourAgent(ByVal agentName As String) As Double
    Dim wsPers As Worksheet, r As Long, lastRow As Long
    Dim nom As String, prenom As String, clef As String
    On Error Resume Next
    Set wsPers = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    If wsPers Is Nothing Then GetHeuresStdJourAgent = 7.6: Exit Function
    lastRow = wsPers.Cells(wsPers.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nom = Trim(SafeCStr(wsPers.Cells(r, 2).Value))
        prenom = Trim(SafeCStr(wsPers.Cells(r, 3).Value))
        clef = nom & "_" & prenom
        If UCase(clef) = UCase(Trim(agentName)) Then
            If IsNumeric(wsPers.Cells(r, 4).Value) And wsPers.Cells(r, 4).Value > 0 Then
                GetHeuresStdJourAgent = CDbl(wsPers.Cells(r, 4).Value)
            Else
                GetHeuresStdJourAgent = 7.6
            End If
            Exit Function
        End If
    Next r
    GetHeuresStdJourAgent = 7.6
End Function

'--------------------------------------------------------------------
' GetPctTempsAgent: recupere le % temps mensuel depuis Personnel
' Colonnes: C29=Janv%, C31=Fev%, C33=Mars% (pas +2 par mois)
' Les valeurs sont des fractions: 0.75 = 75%, 1.0 = 100%
'--------------------------------------------------------------------
Private Function GetPctTempsAgent(ByVal agentName As String, ByVal moisNum As Long) As Double
    Dim wsPers As Worksheet, r As Long, lastRow As Long
    Dim nom As String, prenom As String, clef As String
    Dim colPct As Long, val As Variant
    On Error Resume Next
    Set wsPers = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    If wsPers Is Nothing Then GetPctTempsAgent = 1#: Exit Function
    colPct = 29 + (moisNum - 1) * 2
    lastRow = wsPers.Cells(wsPers.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nom = Trim(SafeCStr(wsPers.Cells(r, 2).Value))
        prenom = Trim(SafeCStr(wsPers.Cells(r, 3).Value))
        clef = nom & "_" & prenom
        If UCase(clef) = UCase(Trim(agentName)) Then
            val = wsPers.Cells(r, colPct).Value
            If IsNumeric(val) And val > 0 Then
                GetPctTempsAgent = CDbl(val)
                If GetPctTempsAgent > 1.5 Then GetPctTempsAgent = GetPctTempsAgent / 100#
            Else
                GetPctTempsAgent = 1#
            End If
            Exit Function
        End If
    Next r
    GetPctTempsAgent = 1#
End Function

'====================================================================
' MACRO PUBLIQUE: GenererBilanHeures
' Cree un onglet recapitulatif annuel des heures par agent
'====================================================================
Public Sub GenererBilanHeures()
    Dim mSheets() As String
    Dim ws As Worksheet, wsBilan As Worksheet
    Dim r As Long, m As Long, rowIdx As Long
    Dim agName As String
    Dim annee As Long
    Dim moisNum As Long
    Dim heuresStd As Double, pctTemps As Double
    Dim hTheo As Double, hPrest As Double
    Dim totalTheo As Double, totalPrest As Double
    Dim headers As Variant
    Dim col As Long
    Dim agentNames As New Collection
    Dim agentRows As Object
    Dim innerDict As Object
    Dim ag As Variant
    Dim baseCol As Long
    Dim d As Double
    Dim totCol As Long
    Dim c2 As Long
    Dim rng As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    mSheets = Split(MONTH_SHEETS, ",")
    annee = Year(Now)

    ' Creer/recreer la feuille Bilan Heures
    DeleteSheetSafe "Bilan Heures"
    Set wsBilan = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsBilan.Name = "Bilan Heures"

    ' En-tetes
    headers = Array("Agent", _
        "Jan Theo", "Jan Prest", "Jan D", _
        "Fev Theo", "Fev Prest", "Fev D", _
        "Mar Theo", "Mar Prest", "Mar D", _
        "Avr Theo", "Avr Prest", "Avr D", _
        "Mai Theo", "Mai Prest", "Mai D", _
        "Jun Theo", "Jun Prest", "Jun D", _
        "Jul Theo", "Jul Prest", "Jul D", _
        "Aou Theo", "Aou Prest", "Aou D", _
        "Sep Theo", "Sep Prest", "Sep D", _
        "Oct Theo", "Oct Prest", "Oct D", _
        "Nov Theo", "Nov Prest", "Nov D", _
        "Dec Theo", "Dec Prest", "Dec D", _
        "TOTAL Theo", "TOTAL Prest", "TOTAL Delta")

    For col = 0 To UBound(headers)
        wsBilan.Cells(1, col + 1).Value = headers(col)
    Next col

    ' Formatage en-tete
    With wsBilan.Range(wsBilan.Cells(1, 1), wsBilan.Cells(1, UBound(headers) + 1))
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With

    ' Collecter les agents depuis le premier mois disponible
    Set agentRows = CreateObject("Scripting.Dictionary")

    For m = 0 To 11
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextMonth2

        For r = FIRST_EMP_ROW To GetLastEmployeeRow(ws)
            If IsError(ws.Cells(r, 1).Value) Then GoTo NextRow2
            agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
            If Len(agName) = 0 Then GoTo NextRow2
            If InStr(agName, "Remplacement") > 0 Then GoTo NextRow2
            If agName = "Us Nuit" Then GoTo NextRow2

            If Not agentRows.Exists(agName) Then
                agentRows.Add agName, CreateObject("Scripting.Dictionary")
                agentNames.Add agName
            End If

            ' Stocker la ligne de cet agent dans ce mois
            Set innerDict = agentRows(agName)
            If Not innerDict.Exists(mSheets(m)) Then
                innerDict.Add mSheets(m), r
            End If
NextRow2:
        Next r
NextMonth2:
        Set ws = Nothing
    Next m

    ' Remplir le bilan
    rowIdx = 2
    For Each ag In agentNames
        agName = CStr(ag)
        wsBilan.Cells(rowIdx, 1).Value = agName

        heuresStd = GetHeuresStdJourAgent(agName)
        If heuresStd <= 0 Then heuresStd = 7.6

        totalTheo = 0#
        totalPrest = 0#

        Set innerDict = agentRows(agName)

        For m = 0 To 11
            moisNum = m + 1
            pctTemps = GetPctTempsAgent(agName, moisNum)
            If pctTemps <= 0 Then pctTemps = 1#

            ' Heures theoriques
            hTheo = HeuresTheoriquesMois(annee, moisNum, pctTemps, heuresStd)

            ' Heures prestees
            If innerDict.Exists(mSheets(m)) Then
                hPrest = HeuresPresteesMois(mSheets(m), CLng(innerDict(mSheets(m))))
            Else
                hPrest = 0#
            End If

            ' Colonnes: 1=Agent, puis groupes de 3 (Theo, Prest, Delta)
            baseCol = 2 + m * 3

            wsBilan.Cells(rowIdx, baseCol).Value = Round(hTheo, 2)
            wsBilan.Cells(rowIdx, baseCol).NumberFormat = "0.00"

            wsBilan.Cells(rowIdx, baseCol + 1).Value = Round(hPrest, 2)
            wsBilan.Cells(rowIdx, baseCol + 1).NumberFormat = "0.00"

            d = Round(hPrest - hTheo, 2)
            wsBilan.Cells(rowIdx, baseCol + 2).Value = d
            wsBilan.Cells(rowIdx, baseCol + 2).NumberFormat = "0.00"

            ' Colorer les deltas
            If d < -5 Then
                wsBilan.Cells(rowIdx, baseCol + 2).Font.Color = RGB(204, 0, 0)
                wsBilan.Cells(rowIdx, baseCol + 2).Font.Bold = True
            ElseIf d > 5 Then
                wsBilan.Cells(rowIdx, baseCol + 2).Font.Color = RGB(0, 128, 0)
                wsBilan.Cells(rowIdx, baseCol + 2).Font.Bold = True
            End If

            totalTheo = totalTheo + hTheo
            totalPrest = totalPrest + hPrest
        Next m

        ' Totaux annuels
        totCol = 2 + 12 * 3 ' Colonne 38

        wsBilan.Cells(rowIdx, totCol).Value = Round(totalTheo, 2)
        wsBilan.Cells(rowIdx, totCol).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol).Font.Bold = True

        wsBilan.Cells(rowIdx, totCol + 1).Value = Round(totalPrest, 2)
        wsBilan.Cells(rowIdx, totCol + 1).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol + 1).Font.Bold = True

        d = Round(totalPrest - totalTheo, 2)
        wsBilan.Cells(rowIdx, totCol + 2).Value = d
        wsBilan.Cells(rowIdx, totCol + 2).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol + 2).Font.Bold = True

        If d < -20 Then
            wsBilan.Cells(rowIdx, totCol + 2).Font.Color = RGB(204, 0, 0)
        ElseIf d > 20 Then
            wsBilan.Cells(rowIdx, totCol + 2).Font.Color = RGB(0, 128, 0)
        End If

        rowIdx = rowIdx + 1
    Next ag

    ' Formatage global
    wsBilan.Columns("A").ColumnWidth = 28
    For c2 = 2 To UBound(headers) + 1
        wsBilan.Columns(c2).ColumnWidth = 8
    Next c2

    ' Bordures
    Set rng = wsBilan.Range(wsBilan.Cells(1, 1), wsBilan.Cells(rowIdx - 1, UBound(headers) + 1))
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlThin
    rng.Font.Name = "Calibri"

    ' Freeze panes
    wsBilan.Range("B2").Select
    ActiveWindow.FreezePanes = True

    wsBilan.Activate

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------
' DeleteSheetSafe: supprime un onglet sans alerte
'--------------------------------------------------------------------
Private Sub DeleteSheetSafe(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

'====================================================================
' FONCTIONS UTILITAIRES PUBLIQUES
' Utilisables depuis d'autres modules ou depuis des formules Excel
'====================================================================

'--------------------------------------------------------------------
' HeuresJourFerie: retourne True si la date est un jour ferie belge
' Utilisable en formule: =HeuresJourFerie(2026, 1, 1) -> True (Nouvel An)
'--------------------------------------------------------------------
Public Function HeuresJourFerie(ByVal annee As Long, _
                                 ByVal mois As Long, _
                                 ByVal jour As Long) As Boolean
    Dim d As Date
    Dim feries As Object

    On Error GoTo ErrExit
    d = DateSerial(annee, mois, jour)
    Set feries = Module_Planning_Core.BuildFeriesBE(annee)
    HeuresJourFerie = feries.Exists(CStr(d))
    Exit Function

ErrExit:
    HeuresJourFerie = False
End Function

'--------------------------------------------------------------------
' HeuresFormatHHMM: convertit heures decimales en format HH:MM
' Ex: 8.5 -> "08:30", 11.75 -> "11:45"
'--------------------------------------------------------------------
Public Function HeuresFormatHHMM(ByVal heuresDecimales As Double) As String
    Dim h As Long, m As Long

    If heuresDecimales < 0 Then
        HeuresFormatHHMM = "-" & HeuresFormatHHMM(Abs(heuresDecimales))
        Exit Function
    End If

    h = Int(heuresDecimales)
    m = Round((heuresDecimales - h) * 60, 0)

    ' Gerer l'arrondi a 60 minutes
    If m >= 60 Then
        h = h + 1
        m = m - 60
    End If

    HeuresFormatHHMM = Format(h, "00") & ":" & Format(m, "00")
End Function
