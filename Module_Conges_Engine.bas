Attribute VB_Name = "Module_Conges_Engine"
'====================================================================
' MODULE CONGES ENGINE - Planning_2026_RUNTIME
' Moteur centralise pour la gestion des conges et absences
'
' FONCTIONNALITES:
'   - Classification centralisee des codes conges (remplace les
'     fonctions dispersees dans Module_SuiviRH et Module_HeuresTravaillees)
'   - Scanning des onglets mensuels pour comptage codes par agent
'   - Calcul et ecriture des soldes dans feuille permanente Soldes_Conges
'   - Journal d'audit dans feuille permanente Historique_Conges
'   - Validation des prises de conge (solde, chevauchement, staffing)
'
' DEPENDANCES:
'   - Feuilles mensuelles: Janv..Dec (Row 6+ = agents, Col 3-33 = jours)
'   - Feuille Personnel: donnees agents (nom, prenom, fonction, %)
'   - Config_Personnel (optionnel): quotas individuels
'   - Module_Planning_Core: BuildFeriesBE() pour jours feries belges
'
' FEUILLES PERMANENTES GEREES:
'   - Soldes_Conges: dashboard temps reel des soldes par agent
'   - Historique_Conges: journal d'audit de tous les mouvements
'
' INSTALLATION:
'   1. Ouvrir Planning_2026_RUNTIME.xlsm
'   2. Alt+F11 > Menu Fichier > Importer > Module_Conges_Engine.bas
'   3. Lancer: Alt+F8 > InitialiserFeuillesConges > Executer
'   4. Puis: Alt+F8 > RecalculerTousSoldes > Executer
'====================================================================

Option Explicit

' ---- CONSTANTES ----
Private Const MONTH_SHEETS As String = "Janv,Fev,Mars,Avril,Mai,Juin,Juil,Aout,Sept,Oct,Nov,Dec"
Private Const FIRST_EMP_ROW As Long = 5
Private Const FIRST_DAY_COL As Long = 3    ' Colonne C = jour 1
Private Const MAX_DAY_COL As Long = 33     ' Colonne AG = jour 31

' Noms des feuilles permanentes
Private Const SHEET_SOLDES As String = "Soldes_Conges"
Private Const SHEET_HISTORIQUE As String = "Historique_Conges"
Private Const SHEET_CONFIG_PERSONNEL As String = "Config_Personnel"
Private Const SHEET_PERSONNEL As String = "Personnel"

' Codes d'absence reconnus (ne comptent PAS comme heures prestees)
Private Const ABSENCE_CODES As String = "CA,EL,ANC,C SOC,DP,CTR,RCT,RV,WE,DECES,RHS,JF"

' Prefixes maladie
Private Const MALADIE_PREFIXES As String = "MAL-,MUT,MAT-,PAT-"

' Staffing minimum par fonction
Private Const MIN_STAFF_INF As Long = 3
Private Const MIN_STAFF_AS As Long = 2

' Colonnes Soldes_Conges
Private Const SC_COL_MATRICULE As Long = 1   ' A
Private Const SC_COL_NOM As Long = 2         ' B
Private Const SC_COL_CA_ACQUIS As Long = 3   ' C
Private Const SC_COL_CA_PRIS As Long = 4     ' D
Private Const SC_COL_CA_SOLDE As Long = 5    ' E
Private Const SC_COL_EL_ACQUIS As Long = 6   ' F
Private Const SC_COL_EL_PRIS As Long = 7     ' G
Private Const SC_COL_EL_SOLDE As Long = 8    ' H
Private Const SC_COL_ANC_ACQUIS As Long = 9  ' I
Private Const SC_COL_ANC_PRIS As Long = 10   ' J
Private Const SC_COL_ANC_SOLDE As Long = 11  ' K
Private Const SC_COL_CSOC_ACQUIS As Long = 12 ' L
Private Const SC_COL_CSOC_PRIS As Long = 13  ' M
Private Const SC_COL_CSOC_SOLDE As Long = 14 ' N
Private Const SC_COL_DP_ACQUIS As Long = 15  ' O
Private Const SC_COL_DP_PRIS As Long = 16    ' P
Private Const SC_COL_DP_SOLDE As Long = 17   ' Q
Private Const SC_COL_CRP_ACQUIS As Long = 18 ' R
Private Const SC_COL_CRP_PRIS As Long = 19   ' S
Private Const SC_COL_CRP_SOLDE As Long = 20  ' T
Private Const SC_COL_DERNIERE_MAJ As Long = 21 ' U

' Colonnes Historique_Conges
Private Const HC_COL_ID As Long = 1          ' A
Private Const HC_COL_DATEHEURE As Long = 2   ' B
Private Const HC_COL_MATRICULE As Long = 3   ' C
Private Const HC_COL_NOM As Long = 4         ' D
Private Const HC_COL_TYPE_CONGE As Long = 5  ' E
Private Const HC_COL_ACTION As Long = 6      ' F
Private Const HC_COL_DATE_DEBUT As Long = 7  ' G
Private Const HC_COL_DATE_FIN As Long = 8    ' H
Private Const HC_COL_NB_JOURS As Long = 9    ' I
Private Const HC_COL_SOLDE_AVANT As Long = 10 ' J
Private Const HC_COL_SOLDE_APRES As Long = 11 ' K
Private Const HC_COL_SOURCE As Long = 12     ' L
Private Const HC_COL_MOIS_PLANNING As Long = 13 ' M
Private Const HC_COL_UTILISATEUR As Long = 14 ' N
Private Const HC_COL_COMMENTAIRE As Long = 15 ' O

'====================================================================
' HELPER: conversion safe cellule -> String
'====================================================================
Private Function SafeCStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeCStr = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeCStr = ""
    Else
        SafeCStr = CStr(v)
    End If
End Function

'====================================================================
' INITIALISATION DES FEUILLES PERMANENTES
'====================================================================

'--------------------------------------------------------------------
' InitialiserFeuillesConges
' Cree les feuilles Soldes_Conges et Historique_Conges si manquantes
' Ne supprime JAMAIS une feuille existante (donnees permanentes)
'--------------------------------------------------------------------
Public Sub InitialiserFeuillesConges()
    Dim wsSoldes As Worksheet
    Dim wsHisto As Worksheet

    Application.ScreenUpdating = False

    ' ---- Creer Soldes_Conges si manquante ----
    On Error Resume Next
    Set wsSoldes = ThisWorkbook.Sheets(SHEET_SOLDES)
    On Error GoTo 0

    If wsSoldes Is Nothing Then
        Set wsSoldes = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSoldes.Name = SHEET_SOLDES
        CreerHeadersSoldes wsSoldes
    End If

    ' ---- Creer Historique_Conges si manquante ----
    On Error Resume Next
    Set wsHisto = ThisWorkbook.Sheets(SHEET_HISTORIQUE)
    On Error GoTo 0

    If wsHisto Is Nothing Then
        Set wsHisto = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsHisto.Name = SHEET_HISTORIQUE
        CreerHeadersHistorique wsHisto
    End If

    Application.ScreenUpdating = True

    MsgBox "Feuilles conges initialisees." & vbCrLf & _
           "  - " & SHEET_SOLDES & " : OK" & vbCrLf & _
           "  - " & SHEET_HISTORIQUE & " : OK", vbInformation, "Module Conges"
End Sub

'--------------------------------------------------------------------
' CreerHeadersSoldes: ecrit les en-tetes de la feuille Soldes_Conges
'--------------------------------------------------------------------
Private Sub CreerHeadersSoldes(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim c As Long

    headers = Array("Matricule", "NomComplet", _
        "CA_Acquis", "CA_Pris", "CA_Solde", _
        "EL_Acquis", "EL_Pris", "EL_Solde", _
        "ANC_Acquis", "ANC_Pris", "ANC_Solde", _
        "CSOC_Acquis", "CSOC_Pris", "CSOC_Solde", _
        "DP_Acquis", "DP_Pris", "DP_Solde", _
        "CRP_Acquis", "CRP_Pris", "CRP_Solde", _
        "DerniereMAJ")

    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    ' Formatage en-tete
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1))
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 28
    Dim colIdx As Long
    For colIdx = 3 To 21
        ws.Columns(colIdx).ColumnWidth = 10
    Next colIdx

    ' Freeze panes
    ws.Range("C2").Select
    ActiveWindow.FreezePanes = True
End Sub

'--------------------------------------------------------------------
' CreerHeadersHistorique: ecrit les en-tetes de Historique_Conges
'--------------------------------------------------------------------
Private Sub CreerHeadersHistorique(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim c As Long

    headers = Array("ID", "DateHeure", "Matricule", "NomComplet", "TypeConge", _
        "Action", "DateDebut", "DateFin", "NbJours", "SoldeAvant", "SoldeApres", _
        "Source", "MoisPlanning", "Utilisateur", "Commentaire")

    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    ' Formatage en-tete
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1))
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 10
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").ColumnWidth = 10
    ws.Columns("L").ColumnWidth = 12
    ws.Columns("M").ColumnWidth = 14
    ws.Columns("N").ColumnWidth = 14
    ws.Columns("O").ColumnWidth = 30

    ' Freeze panes
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

'====================================================================
' CLASSIFICATION CENTRALISEE DES CODES
' Remplace les fonctions dispersees dans Module_SuiviRH (ClassifyCode)
' et Module_HeuresTravaillees (EstCodeAbsence, EstCodeMaladie, EstCodeFerie)
'====================================================================

'--------------------------------------------------------------------
' ClassifierCodeConge
' Classifie un code planning en categorie de conge/absence
'
' @param code   String   Code de la cellule planning
' @return String   Categorie: ca/el/anc/c_soc/dp/ctr/rct/rv/rhs/we/
'                  deces/maladie/crp/ferie/work/other/empty
'--------------------------------------------------------------------
Public Function ClassifierCodeConge(ByVal code As String) As String
    Dim c As String
    c = Trim(UCase(code))

    If Len(c) = 0 Or c = "0" Then
        ClassifierCodeConge = "empty"
        Exit Function
    End If

    Select Case True
        ' --- Conges classiques ---
        Case c = "CA"
            ClassifierCodeConge = "ca"
        Case c = "EL"
            ClassifierCodeConge = "el"
        Case c = "ANC"
            ClassifierCodeConge = "anc"
        Case c = "C SOC"
            ClassifierCodeConge = "c_soc"
        Case c = "DP"
            ClassifierCodeConge = "dp"

        ' --- Compensatoires / recuperation ---
        Case c = "CTR"
            ClassifierCodeConge = "ctr"
        Case c = "RCT"
            ClassifierCodeConge = "rct"
        Case c = "RV"
            ClassifierCodeConge = "rv"
        Case Left(c, 3) = "RHS"
            ClassifierCodeConge = "rhs"

        ' --- Weekend / deces ---
        Case c = "WE"
            ClassifierCodeConge = "we"
        Case c = "DECES"
            ClassifierCodeConge = "deces"
        Case c = "JF"
            ClassifierCodeConge = "ferie"

        ' --- Maladie / maternite / paternite ---
        Case Left(c, 7) = "MAL-GAR", Left(c, 7) = "MAL-MUT"
            ClassifierCodeConge = "maladie"
        Case Left(c, 3) = "MUT"
            ClassifierCodeConge = "maladie"
        Case Left(c, 7) = "MAT-EMP", Left(c, 7) = "MAT-MUT"
            ClassifierCodeConge = "maladie"
        Case Left(c, 7) = "PAT-EMP", Left(c, 7) = "PAT-MUT"
            ClassifierCodeConge = "maladie"

        ' --- Credit / recuperation ---
        Case Left(c, 3) = "CRP"
            ClassifierCodeConge = "crp"

        ' --- Feries (F-xxx, R-xxx) ---
        Case Left(c, 1) = "F" And Len(c) >= 2 And Mid(c, 2, 1) = "-"
            ClassifierCodeConge = "ferie"
        Case Left(c, 1) = "R" And Len(c) >= 2 And Mid(c, 2, 1) = "-"
            ClassifierCodeConge = "ferie"

        ' --- Heures de travail (plages horaires) ---
        Case InStr(c, ":") > 0
            ClassifierCodeConge = "work"
        Case Left(c, 2) = "C "
            ClassifierCodeConge = "work"

        ' --- Non reconnu ---
        Case Else
            ClassifierCodeConge = "other"
    End Select
End Function

'--------------------------------------------------------------------
' EstCodeAbsence
' Verifie si le code est un code d'absence connu
' CA, EL, ANC, C SOC, DP, CTR, RCT, RV, WE, DECES, RHS, JF
'
' @param code   String   Code de la cellule planning
' @return Boolean   True si code d'absence
'--------------------------------------------------------------------
Public Function EstCodeAbsence(ByVal code As String) As Boolean
    Dim c As String
    Dim absCodes() As String
    Dim i As Long

    c = Trim(UCase(code))
    If Len(c) = 0 Then
        EstCodeAbsence = False
        Exit Function
    End If

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
' EstCodeMaladie
' Verifie les prefixes maladie
' MAL-GAR, MAL-MUT, MUT, MAT-EMP, MAT-MUT, PAT-EMP, PAT-MUT
'
' @param code   String   Code de la cellule planning
' @return Boolean   True si code maladie/maternite/paternite
'--------------------------------------------------------------------
Public Function EstCodeMaladie(ByVal code As String) As Boolean
    Dim c As String
    Dim prefixes() As String
    Dim i As Long

    c = Trim(UCase(code))
    If Len(c) = 0 Then
        EstCodeMaladie = False
        Exit Function
    End If

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
' EstCodeFerie
' Detecte les codes feries: F-xxx ou R-xxx (recup ferie)
'
' @param code   String   Code de la cellule planning
' @return Boolean   True si code ferie ou recup ferie
'--------------------------------------------------------------------
Public Function EstCodeFerie(ByVal code As String) As Boolean
    Dim c As String
    c = Trim(UCase(code))

    EstCodeFerie = False
    If Len(c) < 2 Then Exit Function

    If (Left(c, 1) = "F" Or Left(c, 1) = "R") And Mid(c, 2, 1) = "-" Then
        EstCodeFerie = True
    End If
End Function

'====================================================================
' SCANNING DU PLANNING MENSUEL
'====================================================================

'--------------------------------------------------------------------
' ScannerPlanningMensuel
' Parcourt un onglet mois et compte les codes conges par agent
'
' @param nomMois   String   Nom de l'onglet ("Janv", "Fev", etc.)
' @return Object   Dictionary(agentName → Dictionary(typeConge → nbJours))
'--------------------------------------------------------------------
Public Function ScannerPlanningMensuel(ByVal nomMois As String) As Object
    Dim ws As Worksheet
    Dim dictAgents As Object
    Dim dictAgent As Object
    Dim r As Long, c As Long
    Dim agName As String, cellVal As String, cType As String
    Dim numJours As Long

    Set dictAgents = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ScannerPlanningMensuel = dictAgents
        Exit Function
    End If

    numJours = CompterJoursMoisLocal(ws)

    For r = FIRST_EMP_ROW To GetLastEmployeeRowLocal(ws)
        agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
        If Len(agName) = 0 Then GoTo NextRowScan
        If InStr(agName, "Remplacement") > 0 Then GoTo NextRowScan
        If agName = "Us Nuit" Then GoTo NextRowScan

        ' Creer le dictionnaire de l'agent si nouveau
        If Not dictAgents.Exists(agName) Then
            Set dictAgent = CreateObject("Scripting.Dictionary")
            dictAgents.Add agName, dictAgent
        End If
        Set dictAgent = dictAgents(agName)

        ' Scanner les jours
        For c = FIRST_DAY_COL To FIRST_DAY_COL + numJours - 1
            cellVal = Trim(SafeCStr(ws.Cells(r, c).Value))
            If Len(cellVal) = 0 Or cellVal = "0" Then GoTo NextColScan

            cType = ClassifierCodeConge(cellVal)

            ' Ne compter que les types de conge (pas work, empty, other)
            Select Case cType
                Case "ca", "el", "anc", "c_soc", "dp", "ctr", "rct", "rv", _
                     "rhs", "we", "deces", "maladie", "crp", "ferie"
                    If dictAgent.Exists(cType) Then
                        dictAgent(cType) = dictAgent(cType) + 1
                    Else
                        dictAgent.Add cType, 1
                    End If
            End Select
NextColScan:
        Next c
NextRowScan:
    Next r

    Set ScannerPlanningMensuel = dictAgents
End Function

'====================================================================
' CALCUL DES SOLDES
'====================================================================

'--------------------------------------------------------------------
' RecalculerTousSoldes
' Scanne les 12 mois, cumule les codes conges par agent,
' puis met a jour la feuille Soldes_Conges
'--------------------------------------------------------------------
Public Sub RecalculerTousSoldes()
    Dim mSheets() As String
    Dim m As Long
    Dim dictGlobal As Object  ' Dictionary(agent → Dictionary(type → total))
    Dim dictMois As Object
    Dim agKeys As Variant, typeKeys As Variant
    Dim ag As Variant, tp As Variant
    Dim dictAg As Object, dictAgGlobal As Object

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo CleanupRecalc

    Set dictGlobal = CreateObject("Scripting.Dictionary")
    mSheets = Split(MONTH_SHEETS, ",")

    ' Scanner les 12 mois
    For m = 0 To 11
        Set dictMois = ScannerPlanningMensuel(mSheets(m))

        ' Fusionner dans dictGlobal
        agKeys = dictMois.Keys
        Dim idx As Long
        For idx = 0 To dictMois.Count - 1
            Dim agName As String
            agName = agKeys(idx)
            Set dictAg = dictMois(agName)

            If Not dictGlobal.Exists(agName) Then
                Set dictAgGlobal = CreateObject("Scripting.Dictionary")
                dictGlobal.Add agName, dictAgGlobal
            End If
            Set dictAgGlobal = dictGlobal(agName)

            typeKeys = dictAg.Keys
            Dim idx2 As Long
            For idx2 = 0 To dictAg.Count - 1
                Dim tpName As String
                tpName = typeKeys(idx2)
                If dictAgGlobal.Exists(tpName) Then
                    dictAgGlobal(tpName) = dictAgGlobal(tpName) + dictAg(tpName)
                Else
                    dictAgGlobal.Add tpName, dictAg(tpName)
                End If
            Next idx2
        Next idx
    Next m

    ' Ecrire les resultats
    EcrireSoldesConges dictGlobal

CleanupRecalc:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Erreur recalcul soldes: " & Err.Description, vbCritical, "Module Conges"
    End If
End Sub

'--------------------------------------------------------------------
' RecalculerSoldesMois
' Recalcule les soldes en rescannant tous les mois
' (un seul mois ne suffit pas car les soldes sont cumulatifs)
'
' @param nomMois   String   Nom du mois qui a change (pour info)
'--------------------------------------------------------------------
Public Sub RecalculerSoldesMois(ByVal nomMois As String)
    ' Le recalcul est toujours global car les soldes sont annuels
    ' Le parametre nomMois sert de contexte pour le log
    RecalculerTousSoldes
End Sub

'--------------------------------------------------------------------
' CompterCongesParAgent
' Compte les jours de conge d'un type specifique pour un agent
' sur une plage de mois
'
' @param nom         String   Nom de l'agent (format "Nom_Prenom")
' @param typeConge   String   Type de conge ("ca", "el", "anc", etc.)
' @param moisDeb     Long     Mois de debut (1-12), defaut 1
' @param moisFin     Long     Mois de fin (1-12), defaut 12
' @return Double   Nombre de jours
'--------------------------------------------------------------------
Public Function CompterCongesParAgent(ByVal nom As String, _
                                       ByVal typeConge As String, _
                                       Optional ByVal moisDeb As Long = 1, _
                                       Optional ByVal moisFin As Long = 12) As Double
    Dim mSheets() As String
    Dim m As Long
    Dim dictMois As Object
    Dim dictAg As Object
    Dim total As Double
    Dim tc As String

    mSheets = Split(MONTH_SHEETS, ",")
    tc = LCase(Trim(typeConge))
    total = 0

    ' Valider les bornes
    If moisDeb < 1 Then moisDeb = 1
    If moisFin > 12 Then moisFin = 12
    If moisDeb > moisFin Then
        CompterCongesParAgent = 0
        Exit Function
    End If

    For m = moisDeb - 1 To moisFin - 1
        Set dictMois = ScannerPlanningMensuel(mSheets(m))
        If dictMois.Exists(nom) Then
            Set dictAg = dictMois(nom)
            If dictAg.Exists(tc) Then
                total = total + dictAg(tc)
            End If
        End If
    Next m

    CompterCongesParAgent = total
End Function

'--------------------------------------------------------------------
' GetSoldeAgent
' Lit le solde actuel d'un type de conge depuis la feuille Soldes_Conges
'
' @param nom         String   Nom complet de l'agent
' @param typeConge   String   Type: "ca", "el", "anc", "c_soc", "dp", "crp"
' @return Double   Solde restant (peut etre negatif)
'--------------------------------------------------------------------
Public Function GetSoldeAgent(ByVal nom As String, _
                               ByVal typeConge As String) As Double
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim colSolde As Long
    Dim agName As String

    GetSoldeAgent = 0

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SOLDES)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    colSolde = GetSoldeColumnForType(typeConge)
    If colSolde = 0 Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, SC_COL_NOM).End(xlUp).Row
    For r = 2 To lastRow
        agName = Trim(SafeCStr(ws.Cells(r, SC_COL_NOM).Value))
        If UCase(agName) = UCase(Trim(nom)) Then
            If IsNumeric(ws.Cells(r, colSolde).Value) Then
                GetSoldeAgent = CDbl(ws.Cells(r, colSolde).Value)
            End If
            Exit Function
        End If
    Next r
End Function

'--------------------------------------------------------------------
' GetQuotaAgent
' Lit le quota annuel d'un type de conge pour un agent
' Source prioritaire: Config_Personnel, fallback: quotas hardcodes
'
' @param nom         String   Nom complet de l'agent
' @param typeConge   String   Type: "ca", "el", "anc", "c_soc", "dp", "crp"
' @return Double   Quota annuel
'--------------------------------------------------------------------
Public Function GetQuotaAgent(ByVal nom As String, _
                               ByVal typeConge As String) As Double
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim colQuota As Long
    Dim agName As String
    Dim tc As String

    GetQuotaAgent = 0
    tc = LCase(Trim(typeConge))

    ' Essayer Config_Personnel d'abord
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG_PERSONNEL)
    On Error GoTo 0

    If Not ws Is Nothing Then
        colQuota = GetConfigPersonnelQuotaCol(tc)
        If colQuota > 0 Then
            lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            For r = 2 To lastRow
                ' Config_Personnel: col B = Nom, col C = Prenom
                Dim nomComplet As String
                nomComplet = Trim(SafeCStr(ws.Cells(r, 2).Value)) & "_" & _
                             Trim(SafeCStr(ws.Cells(r, 3).Value))
                If UCase(nomComplet) = UCase(Trim(nom)) Then
                    If IsNumeric(ws.Cells(r, colQuota).Value) Then
                        GetQuotaAgent = CDbl(ws.Cells(r, colQuota).Value)
                        Exit Function
                    End If
                End If
            Next r
        End If
    End If

    ' Fallback: quotas hardcodes depuis Module_SuiviRH
    GetQuotaAgent = GetQuotaHardcode(nom, tc)
End Function

'--------------------------------------------------------------------
' GetConfigPersonnelQuotaCol: mappe type conge -> colonne Config_Personnel
' Config_Personnel: J=QuotaCA, K=QuotaEL, L=QuotaANC, M=QuotaCSoc,
'                   N=QuotaDP, O=QuotaCRP
'--------------------------------------------------------------------
Private Function GetConfigPersonnelQuotaCol(ByVal typeConge As String) As Long
    Select Case LCase(Trim(typeConge))
        Case "ca": GetConfigPersonnelQuotaCol = 10     ' J
        Case "el": GetConfigPersonnelQuotaCol = 11     ' K
        Case "anc": GetConfigPersonnelQuotaCol = 12    ' L
        Case "c_soc": GetConfigPersonnelQuotaCol = 13  ' M
        Case "dp": GetConfigPersonnelQuotaCol = 14     ' N
        Case "crp": GetConfigPersonnelQuotaCol = 15    ' O
        Case Else: GetConfigPersonnelQuotaCol = 0
    End Select
End Function

'--------------------------------------------------------------------
' GetQuotaHardcode: fallback quotas hardcodes
' Source: Module_SuiviRH.GetQuotas (25 agents)
'--------------------------------------------------------------------
Private Function GetQuotaHardcode(ByVal nom As String, ByVal typeConge As String) As Double
    Dim quotas As Object
    Dim key As String

    Set quotas = CreateObject("Scripting.Dictionary")

    ' Format: "NOM_PRENOM" -> Array(CA, EL, ANC, CSOC, DP, CRP)
    quotas.Add "HERMANN_CLAUDE", Array(24, 5, 4, 2, 1, 0)
    quotas.Add "BEN ABDELKADER_YAHYA", Array(24, 3, 3, 2, 0, 0)
    quotas.Add "BOURGEOIS_AURORE", Array(24, 5, 4, 2, 1, 8.35)
    quotas.Add "OURTIOUALOUS_NAA" & ChrW(206) & "MA", Array(24, 8, 1, 2, 1, 0)
    quotas.Add "BOZIC_JACQUELINE", Array(24, 5, 0, 2, 0, 0)
    quotas.Add "YOUSSOUF_ROUKKIAT", Array(24, 2, 3, 2, 2, 0)
    quotas.Add "WIELEMANS_JENNELIE", Array(24, 5, 2, 2, 0, 0)
    quotas.Add "EL GHARBAOUI_SH" & ChrW(201) & "RAZADE", Array(24, 5, 3, 2, 0, 0)
    quotas.Add "MUPIKA MANGA_CAROLINE", Array(24, 4, 3, 2, 2, 0)
    quotas.Add "ULPAT_VICTOR", Array(24, 5, 1, 2, 0, 0)
    quotas.Add "HAOURIQUI_MOHAMED", Array(24, 5, 1, 2, 0, 0)
    quotas.Add "VORST_JULIE", Array(24, 5, 0, 2, 0, 0)
    quotas.Add "DIALLO_MAMADOU", Array(24, 5, 0, 2, 0, 0)
    quotas.Add "DELA VEGA_EDELYN", Array(24, 5, 1, 2, 0, 7.96)
    quotas.Add "OUSROUT_SALMA", Array(24, 5, 0, 2, 0, 0)
    quotas.Add "MUTOMBO ILUNGA_FRANCIS", Array(24, 1, 1, 2, 0, 0)
    quotas.Add "BOSSAERT_MARION", Array(24, 3, 0, 2, 0, 0)
    quotas.Add "DE BUS_ANJA", Array(24, 5, 0, 2, 0, 0)
    quotas.Add "ADZOGBLE_CHARLES", Array(24, 5, 3, 2, 0, 0)
    quotas.Add "NANA CHAMBA_HENRI", Array(24, 2, 1, 2, 0, 0)
    quotas.Add "DE SMEDT_SABRINA", Array(24, 3, 0, 2, 0, 0)
    quotas.Add "ALAHYANE_ZAHRA", Array(24, 5, 0, 2, 0, 0)
    quotas.Add "UWERA_LAETITIA", Array(24, 5, 2, 2, 3, 0)
    quotas.Add "NAYITURIKIV_V" & ChrW(201) & "R" & ChrW(200) & "NE", Array(24, 4, 1, 2, 1, 0)
    quotas.Add "RAMACK_SYLVIE", Array(24, 5, 7, 2, 3, 10)

    key = UCase(Trim(nom))
    GetQuotaHardcode = 0

    If Not quotas.Exists(key) Then
        ' Defauts generiques si agent inconnu
        Select Case typeConge
            Case "ca": GetQuotaHardcode = 24
            Case "el": GetQuotaHardcode = 5
            Case "c_soc": GetQuotaHardcode = 2
            Case Else: GetQuotaHardcode = 0
        End Select
        Exit Function
    End If

    Dim arr As Variant
    arr = quotas(key)

    Select Case typeConge
        Case "ca": GetQuotaHardcode = arr(0)
        Case "el": GetQuotaHardcode = arr(1)
        Case "anc": GetQuotaHardcode = arr(2)
        Case "c_soc": GetQuotaHardcode = arr(3)
        Case "dp": GetQuotaHardcode = arr(4)
        Case "crp": GetQuotaHardcode = arr(5)
        Case Else: GetQuotaHardcode = 0
    End Select
End Function

'--------------------------------------------------------------------
' GetSoldeColumnForType: mappe type conge -> colonne solde dans Soldes_Conges
'--------------------------------------------------------------------
Private Function GetSoldeColumnForType(ByVal typeConge As String) As Long
    Select Case LCase(Trim(typeConge))
        Case "ca": GetSoldeColumnForType = SC_COL_CA_SOLDE
        Case "el": GetSoldeColumnForType = SC_COL_EL_SOLDE
        Case "anc": GetSoldeColumnForType = SC_COL_ANC_SOLDE
        Case "c_soc": GetSoldeColumnForType = SC_COL_CSOC_SOLDE
        Case "dp": GetSoldeColumnForType = SC_COL_DP_SOLDE
        Case "crp": GetSoldeColumnForType = SC_COL_CRP_SOLDE
        Case Else: GetSoldeColumnForType = 0
    End Select
End Function

'====================================================================
' ECRITURE SOLDES ET HISTORIQUE
'====================================================================

'--------------------------------------------------------------------
' EcrireSoldesConges
' Ecrit les soldes cumules dans la feuille Soldes_Conges
' Met a jour les lignes existantes ou en cree de nouvelles
'
' @param dictGlobal   Object   Dictionary(agent → Dictionary(typeConge → nbJours))
'--------------------------------------------------------------------
Public Sub EcrireSoldesConges(ByVal dictGlobal As Object)
    Dim ws As Worksheet
    Dim r As Long, nextRow As Long
    Dim agKeys As Variant
    Dim idx As Long
    Dim agName As String
    Dim dictAg As Object
    Dim existingRow As Long
    Dim quotaCA As Double, quotaEL As Double, quotaANC As Double
    Dim quotaCSoc As Double, quotaDP As Double, quotaCRP As Double
    Dim prisCA As Double, prisEL As Double, prisANC As Double
    Dim prisCSoc As Double, prisDP As Double, prisCRP As Double

    ' S'assurer que la feuille existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SOLDES)
    On Error GoTo 0
    If ws Is Nothing Then
        InitialiserFeuillesConges
        Set ws = ThisWorkbook.Sheets(SHEET_SOLDES)
        If ws Is Nothing Then Exit Sub
    End If

    ' Desactiver events
    Application.EnableEvents = False

    agKeys = dictGlobal.Keys
    For idx = 0 To dictGlobal.Count - 1
        agName = agKeys(idx)
        Set dictAg = dictGlobal(agName)

        ' Lire les jours pris
        prisCA = 0: prisEL = 0: prisANC = 0
        prisCSoc = 0: prisDP = 0: prisCRP = 0

        If dictAg.Exists("ca") Then prisCA = dictAg("ca")
        If dictAg.Exists("el") Then prisEL = dictAg("el")
        If dictAg.Exists("anc") Then prisANC = dictAg("anc")
        If dictAg.Exists("c_soc") Then prisCSoc = dictAg("c_soc")
        If dictAg.Exists("dp") Then prisDP = dictAg("dp")
        If dictAg.Exists("crp") Then prisCRP = dictAg("crp")

        ' Lire les quotas
        quotaCA = GetQuotaAgent(agName, "ca")
        quotaEL = GetQuotaAgent(agName, "el")
        quotaANC = GetQuotaAgent(agName, "anc")
        quotaCSoc = GetQuotaAgent(agName, "c_soc")
        quotaDP = GetQuotaAgent(agName, "dp")
        quotaCRP = GetQuotaAgent(agName, "crp")

        ' Chercher si l'agent existe deja
        existingRow = TrouverLigneAgent(ws, agName)
        If existingRow = 0 Then
            ' Nouvelle ligne
            existingRow = ws.Cells(ws.Rows.Count, SC_COL_NOM).End(xlUp).Row + 1
            If existingRow < 2 Then existingRow = 2
            ws.Cells(existingRow, SC_COL_MATRICULE).Value = GetMatriculeAgent(agName)
            ws.Cells(existingRow, SC_COL_NOM).Value = agName
        End If

        r = existingRow

        ' Ecrire les valeurs
        ws.Cells(r, SC_COL_CA_ACQUIS).Value = quotaCA
        ws.Cells(r, SC_COL_CA_PRIS).Value = prisCA
        ws.Cells(r, SC_COL_CA_SOLDE).Value = quotaCA - prisCA

        ws.Cells(r, SC_COL_EL_ACQUIS).Value = quotaEL
        ws.Cells(r, SC_COL_EL_PRIS).Value = prisEL
        ws.Cells(r, SC_COL_EL_SOLDE).Value = quotaEL - prisEL

        ws.Cells(r, SC_COL_ANC_ACQUIS).Value = quotaANC
        ws.Cells(r, SC_COL_ANC_PRIS).Value = prisANC
        ws.Cells(r, SC_COL_ANC_SOLDE).Value = quotaANC - prisANC

        ws.Cells(r, SC_COL_CSOC_ACQUIS).Value = quotaCSoc
        ws.Cells(r, SC_COL_CSOC_PRIS).Value = prisCSoc
        ws.Cells(r, SC_COL_CSOC_SOLDE).Value = quotaCSoc - prisCSoc

        ws.Cells(r, SC_COL_DP_ACQUIS).Value = quotaDP
        ws.Cells(r, SC_COL_DP_PRIS).Value = prisDP
        ws.Cells(r, SC_COL_DP_SOLDE).Value = quotaDP - prisDP

        ws.Cells(r, SC_COL_CRP_ACQUIS).Value = quotaCRP
        ws.Cells(r, SC_COL_CRP_PRIS).Value = prisCRP
        ws.Cells(r, SC_COL_CRP_SOLDE).Value = quotaCRP - prisCRP

        ws.Cells(r, SC_COL_DERNIERE_MAJ).Value = Now
        ws.Cells(r, SC_COL_DERNIERE_MAJ).NumberFormat = "dd/mm/yyyy hh:mm"

        ' Colorer les soldes negatifs en rouge
        Dim soldeCol As Variant
        For Each soldeCol In Array(SC_COL_CA_SOLDE, SC_COL_EL_SOLDE, SC_COL_ANC_SOLDE, _
                                    SC_COL_CSOC_SOLDE, SC_COL_DP_SOLDE, SC_COL_CRP_SOLDE)
            If ws.Cells(r, CLng(soldeCol)).Value < 0 Then
                ws.Cells(r, CLng(soldeCol)).Font.Color = RGB(204, 0, 0)
                ws.Cells(r, CLng(soldeCol)).Font.Bold = True
                ws.Cells(r, CLng(soldeCol)).Interior.Color = RGB(255, 199, 206)
            Else
                ws.Cells(r, CLng(soldeCol)).Font.Color = RGB(0, 102, 0)
                ws.Cells(r, CLng(soldeCol)).Font.Bold = False
                ws.Cells(r, CLng(soldeCol)).Interior.ColorIndex = xlNone
            End If
        Next

    Next idx

    Application.EnableEvents = True
End Sub

'--------------------------------------------------------------------
' EcrireHistorique
' Ajoute une ligne dans le journal d'audit Historique_Conges
'
' @param matricule    Variant   Matricule de l'agent
' @param nom          Variant   Nom complet
' @param typeConge    Variant   Type de conge
' @param action       Variant   PRISE / ANNULATION / AJUSTEMENT
' @param dateDeb      Variant   Date de debut
' @param dateFin      Variant   Date de fin
' @param nbJours      Variant   Nombre de jours
' @param soldeAvant   Variant   Solde avant l'operation
' @param soldeApres   Variant   Solde apres l'operation
' @param source       Variant   Source: "PLANNING" / "FORMULAIRE" / "RECALCUL"
' @param commentaire  Variant   Commentaire libre
'--------------------------------------------------------------------
Public Sub EcrireHistorique(ByVal matricule As Variant, _
                             ByVal nom As Variant, _
                             ByVal typeConge As Variant, _
                             ByVal action As Variant, _
                             ByVal dateDeb As Variant, _
                             ByVal dateFin As Variant, _
                             ByVal nbJours As Variant, _
                             ByVal soldeAvant As Variant, _
                             ByVal soldeApres As Variant, _
                             ByVal source As Variant, _
                             ByVal commentaire As Variant)
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim nextID As Long
    Dim moisPlanning As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_HISTORIQUE)
    On Error GoTo 0
    If ws Is Nothing Then
        InitialiserFeuillesConges
        Set ws = ThisWorkbook.Sheets(SHEET_HISTORIQUE)
        If ws Is Nothing Then Exit Sub
    End If

    ' Trouver la prochaine ligne vide
    nextRow = ws.Cells(ws.Rows.Count, HC_COL_ID).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ' Calculer le prochain ID
    If nextRow = 2 Then
        nextID = 1
    Else
        If IsNumeric(ws.Cells(nextRow - 1, HC_COL_ID).Value) Then
            nextID = CLng(ws.Cells(nextRow - 1, HC_COL_ID).Value) + 1
        Else
            nextID = nextRow - 1
        End If
    End If

    ' Determiner le mois du planning
    If IsDate(dateDeb) Then
        moisPlanning = GetNomMoisFromNum(Month(CDate(dateDeb)))
    Else
        moisPlanning = ""
    End If

    Application.EnableEvents = False

    ws.Cells(nextRow, HC_COL_ID).Value = nextID
    ws.Cells(nextRow, HC_COL_DATEHEURE).Value = Now
    ws.Cells(nextRow, HC_COL_DATEHEURE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(nextRow, HC_COL_MATRICULE).Value = matricule
    ws.Cells(nextRow, HC_COL_NOM).Value = nom
    ws.Cells(nextRow, HC_COL_TYPE_CONGE).Value = typeConge
    ws.Cells(nextRow, HC_COL_ACTION).Value = action
    ws.Cells(nextRow, HC_COL_DATE_DEBUT).Value = dateDeb
    ws.Cells(nextRow, HC_COL_DATE_FIN).Value = dateFin
    ws.Cells(nextRow, HC_COL_NB_JOURS).Value = nbJours
    ws.Cells(nextRow, HC_COL_SOLDE_AVANT).Value = soldeAvant
    ws.Cells(nextRow, HC_COL_SOLDE_APRES).Value = soldeApres
    ws.Cells(nextRow, HC_COL_SOURCE).Value = source
    ws.Cells(nextRow, HC_COL_MOIS_PLANNING).Value = moisPlanning
    ws.Cells(nextRow, HC_COL_UTILISATEUR).Value = Application.UserName
    ws.Cells(nextRow, HC_COL_COMMENTAIRE).Value = commentaire

    ' Formatage de la ligne
    ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 15)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 15)).Borders.Weight = xlThin

    Application.EnableEvents = True
End Sub

'====================================================================
' VALIDATION DE PRISE DE CONGE
'====================================================================

'--------------------------------------------------------------------
' ValiderPriseConge
' Verifie toutes les conditions pour une prise de conge valide
'
' @param nom         String   Nom complet de l'agent
' @param typeConge   String   Type de conge ("ca", "el", "anc", etc.)
' @param dateDeb     Date     Date de debut
' @param dateFin     Date     Date de fin
' @param msg         String   (ByRef) Message d'erreur si validation echoue
' @return Boolean   True si la prise de conge est valide
'--------------------------------------------------------------------
Public Function ValiderPriseConge(ByVal nom As String, _
                                   ByVal typeConge As String, _
                                   ByVal dateDeb As Date, _
                                   ByVal dateFin As Date, _
                                   ByRef msg As String) As Boolean
    Dim solde As Double
    Dim nbJours As Long
    Dim tc As String

    msg = ""
    ValiderPriseConge = False
    tc = LCase(Trim(typeConge))

    ' ---- 1. Verifier que les dates sont coherentes ----
    If dateFin < dateDeb Then
        msg = "La date de fin ne peut pas etre anterieure a la date de debut."
        Exit Function
    End If

    ' ---- 2. Verifier que la date n'est pas dans le passe ----
    If dateDeb < Date Then
        msg = "Impossible de poser un conge dans le passe (debut: " & _
              Format(dateDeb, "dd/mm/yyyy") & ")."
        Exit Function
    End If

    ' ---- 3. Calculer le nombre de jours ouvrables ----
    nbJours = CalculerJoursOuvrablesPeriode(dateDeb, dateFin)
    If nbJours = 0 Then
        msg = "La periode selectionnee ne contient aucun jour ouvrable."
        Exit Function
    End If

    ' ---- 4. Verifier le solde (sauf pour maladie, we, ferie, ctr) ----
    Select Case tc
        Case "maladie", "we", "ferie", "ctr", "rct", "deces"
            ' Pas de verification de solde pour ces types
        Case Else
            solde = GetSoldeAgent(nom, tc)
            If solde < nbJours Then
                msg = "Solde insuffisant pour " & UCase(tc) & ": " & _
                      "solde=" & solde & ", demande=" & nbJours & " jours."
                Exit Function
            End If
    End Select

    ' ---- 5. Verifier le chevauchement avec des conges existants ----
    If ExisteChevauchementConge(nom, dateDeb, dateFin) Then
        msg = "Un conge existe deja sur cette periode pour " & nom & "."
        Exit Function
    End If

    ' ---- 6. Verifier le staffing minimum ----
    Dim staffMsg As String
    If Not VerifierStaffingMinimum(nom, dateDeb, dateFin, staffMsg) Then
        msg = staffMsg
        Exit Function
    End If

    ' Tout OK
    ValiderPriseConge = True
End Function

'--------------------------------------------------------------------
' CalculerJoursOuvrablesPeriode
' Compte les jours ouvrables (Lun-Ven) entre deux dates
' en excluant les jours feries belges
'
' @param dateDeb   Date   Date de debut
' @param dateFin   Date   Date de fin
' @return Long   Nombre de jours ouvrables
'--------------------------------------------------------------------
Private Function CalculerJoursOuvrablesPeriode(ByVal dateDeb As Date, _
                                                ByVal dateFin As Date) As Long
    Dim d As Date
    Dim compteur As Long
    Dim feries As Object
    Dim annee As Long

    compteur = 0
    annee = Year(dateDeb)

    On Error Resume Next
    Set feries = Module_Planning_Core.BuildFeriesBE(annee)
    On Error GoTo 0

    If feries Is Nothing Then
        Set feries = CreateObject("Scripting.Dictionary")
    End If

    ' Si la periode chevauche deux annees, charger les feries de l'annee suivante
    Dim feries2 As Object
    If Year(dateFin) > annee Then
        On Error Resume Next
        Set feries2 = Module_Planning_Core.BuildFeriesBE(Year(dateFin))
        On Error GoTo 0
        If Not feries2 Is Nothing Then
            Dim k As Variant
            For Each k In feries2.Keys
                If Not feries.Exists(k) Then feries.Add k, feries2(k)
            Next k
        End If
    End If

    For d = dateDeb To dateFin
        ' Exclure samedi (7) et dimanche (1) avec vbMonday
        If Weekday(d, vbMonday) <= 5 Then
            If Not feries.Exists(CStr(d)) Then
                compteur = compteur + 1
            End If
        End If
    Next d

    CalculerJoursOuvrablesPeriode = compteur
End Function

'--------------------------------------------------------------------
' ExisteChevauchementConge
' Verifie si un agent a deja un conge sur la periode donnee
' en scannant les onglets mensuels concernes
'--------------------------------------------------------------------
Private Function ExisteChevauchementConge(ByVal nom As String, _
                                           ByVal dateDeb As Date, _
                                           ByVal dateFin As Date) As Boolean
    Dim mSheets() As String
    Dim ws As Worksheet
    Dim moisDeb As Long, moisFin As Long
    Dim m As Long, r As Long, c As Long
    Dim agRow As Long
    Dim numJours As Long
    Dim jourDuMois As Long
    Dim cellVal As String, cType As String
    Dim dateCell As Date

    mSheets = Split(MONTH_SHEETS, ",")
    moisDeb = Month(dateDeb)
    moisFin = Month(dateFin)
    ExisteChevauchementConge = False

    For m = moisDeb To moisFin
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m - 1))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextMonthChev

        ' Trouver la ligne de l'agent
        agRow = TrouverLigneAgentDansMois(ws, nom)
        If agRow = 0 Then GoTo NextMonthChev

        numJours = CompterJoursMoisLocal(ws)

        ' Scanner les jours de la periode dans ce mois
        For c = FIRST_DAY_COL To FIRST_DAY_COL + numJours - 1
            jourDuMois = c - FIRST_DAY_COL + 1
            On Error Resume Next
            dateCell = DateSerial(Year(dateDeb), m, jourDuMois)
            On Error GoTo 0

            If dateCell >= dateDeb And dateCell <= dateFin Then
                cellVal = Trim(SafeCStr(ws.Cells(agRow, c).Value))
                If Len(cellVal) > 0 And cellVal <> "0" Then
                    cType = ClassifierCodeConge(cellVal)
                    If cType <> "work" And cType <> "empty" And cType <> "other" And cType <> "we" Then
                        ExisteChevauchementConge = True
                        Exit Function
                    End If
                End If
            End If
        Next c
NextMonthChev:
        Set ws = Nothing
    Next m
End Function

'--------------------------------------------------------------------
' VerifierStaffingMinimum
' Verifie que le staffing minimum est respecte sur la periode
' Regle: minimum 3 INF et 2 AS presents par jour
'--------------------------------------------------------------------
Private Function VerifierStaffingMinimum(ByVal nom As String, _
                                          ByVal dateDeb As Date, _
                                          ByVal dateFin As Date, _
                                          ByRef msg As String) As Boolean
    Dim mSheets() As String
    Dim ws As Worksheet
    Dim wsPers As Worksheet
    Dim m As Long, r As Long, c As Long
    Dim jourDuMois As Long
    Dim dateCell As Date
    Dim numJours As Long
    Dim agName As String, cellVal As String, cType As String
    Dim fonctionAgent As String
    Dim countINF As Long, countAS As Long
    Dim lastR As Long

    mSheets = Split(MONTH_SHEETS, ",")
    VerifierStaffingMinimum = True

    ' Determiner la fonction de l'agent qui veut prendre conge
    fonctionAgent = GetFonctionAgent(nom)

    For m = Month(dateDeb) To Month(dateFin)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m - 1))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextMonthStaff

        numJours = CompterJoursMoisLocal(ws)
        lastR = GetLastEmployeeRowLocal(ws)

        For c = FIRST_DAY_COL To FIRST_DAY_COL + numJours - 1
            jourDuMois = c - FIRST_DAY_COL + 1
            On Error Resume Next
            dateCell = DateSerial(Year(dateDeb), m, jourDuMois)
            On Error GoTo 0

            If dateCell >= dateDeb And dateCell <= dateFin Then
                ' Weekday check: on ne verifie le staffing que les jours ouvrables
                If Weekday(dateCell, vbMonday) <= 5 Then
                    countINF = 0
                    countAS = 0

                    For r = FIRST_EMP_ROW To lastR
                        agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
                        If Len(agName) = 0 Then GoTo NextRowStaff
                        If InStr(agName, "Remplacement") > 0 Then GoTo NextRowStaff
                        If agName = "Us Nuit" Then GoTo NextRowStaff

                        ' Exclure l'agent qui veut prendre conge
                        If UCase(agName) = UCase(nom) Then GoTo NextRowStaff

                        cellVal = Trim(SafeCStr(ws.Cells(r, c).Value))
                        If Len(cellVal) > 0 Then
                            cType = ClassifierCodeConge(cellVal)
                            ' L'agent est present si son code est "work"
                            If cType = "work" Then
                                Dim fonctionAutre As String
                                fonctionAutre = GetFonctionAgent(agName)
                                If fonctionAutre = "INF" Then
                                    countINF = countINF + 1
                                ElseIf fonctionAutre = "AS" Then
                                    countAS = countAS + 1
                                End If
                            End If
                        End If
NextRowStaff:
                    Next r

                    ' Verifier les minimums
                    If fonctionAgent = "INF" And countINF < MIN_STAFF_INF Then
                        msg = "Staffing insuffisant le " & Format(dateCell, "dd/mm/yyyy") & _
                              ": seulement " & countINF & " INF presents (minimum " & MIN_STAFF_INF & ")."
                        VerifierStaffingMinimum = False
                        Exit Function
                    End If
                    If fonctionAgent = "AS" And countAS < MIN_STAFF_AS Then
                        msg = "Staffing insuffisant le " & Format(dateCell, "dd/mm/yyyy") & _
                              ": seulement " & countAS & " AS presents (minimum " & MIN_STAFF_AS & ")."
                        VerifierStaffingMinimum = False
                        Exit Function
                    End If
                End If
            End If
        Next c
NextMonthStaff:
        Set ws = Nothing
    Next m
End Function

'====================================================================
' FONCTIONS PRIVEES DE LECTURE
'====================================================================

'--------------------------------------------------------------------
' TrouverLigneAgent: cherche la ligne d'un agent dans Soldes_Conges
' Retourne 0 si non trouve
'--------------------------------------------------------------------
Private Function TrouverLigneAgent(ByVal ws As Worksheet, _
                                    ByVal nom As String) As Long
    Dim r As Long, lastRow As Long
    Dim agName As String

    TrouverLigneAgent = 0
    lastRow = ws.Cells(ws.Rows.Count, SC_COL_NOM).End(xlUp).Row

    For r = 2 To lastRow
        agName = Trim(SafeCStr(ws.Cells(r, SC_COL_NOM).Value))
        If UCase(agName) = UCase(Trim(nom)) Then
            TrouverLigneAgent = r
            Exit Function
        End If
    Next r
End Function

'--------------------------------------------------------------------
' TrouverLigneAgentDansMois: cherche la ligne d'un agent dans un onglet mois
'--------------------------------------------------------------------
Private Function TrouverLigneAgentDansMois(ByVal ws As Worksheet, _
                                            ByVal nom As String) As Long
    Dim r As Long
    Dim agName As String

    TrouverLigneAgentDansMois = 0
    For r = FIRST_EMP_ROW To GetLastEmployeeRowLocal(ws)
        agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
        If UCase(agName) = UCase(Trim(nom)) Then
            TrouverLigneAgentDansMois = r
            Exit Function
        End If
    Next r
End Function

'--------------------------------------------------------------------
' CompterJoursMoisLocal: compte les jours d'un mois
'--------------------------------------------------------------------
Private Function CompterJoursMoisLocal(ByVal ws As Worksheet) As Long
    Dim col As Long
    Dim compteur As Long

    compteur = 0
    For col = FIRST_DAY_COL To MAX_DAY_COL
        If IsNumeric(ws.Cells(4, col).Value) And _
           Len(Trim(SafeCStr(ws.Cells(4, col).Value))) > 0 Then
            compteur = compteur + 1
        Else
            If compteur > 0 Then Exit For
        End If
    Next col

    If compteur < 28 Then compteur = 31
    If compteur > 31 Then compteur = 31

    CompterJoursMoisLocal = compteur
End Function

'--------------------------------------------------------------------
' GetLastEmployeeRowLocal: trouve la derniere ligne agent
'--------------------------------------------------------------------
Private Function GetLastEmployeeRowLocal(ByVal ws As Worksheet) As Long
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

    GetLastEmployeeRowLocal = lastFound
End Function

'--------------------------------------------------------------------
' GetMatriculeAgent: recupere le matricule depuis la feuille Personnel
' Retourne le nom comme fallback si non trouve
'--------------------------------------------------------------------
Private Function GetMatriculeAgent(ByVal nom As String) As String
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim nomComplet As String

    GetMatriculeAgent = nom ' Fallback

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_PERSONNEL)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nomComplet = Trim(SafeCStr(ws.Cells(r, 2).Value)) & "_" & _
                     Trim(SafeCStr(ws.Cells(r, 3).Value))
        If UCase(nomComplet) = UCase(Trim(nom)) Then
            If Len(Trim(SafeCStr(ws.Cells(r, 1).Value))) > 0 Then
                GetMatriculeAgent = Trim(SafeCStr(ws.Cells(r, 1).Value))
            End If
            Exit Function
        End If
    Next r
End Function

'--------------------------------------------------------------------
' GetFonctionAgent: recupere la fonction (INF, AS, CEFA) depuis Personnel
'--------------------------------------------------------------------
Private Function GetFonctionAgent(ByVal nom As String) As String
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim nomComplet As String

    GetFonctionAgent = "" ' Inconnu

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_PERSONNEL)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nomComplet = Trim(SafeCStr(ws.Cells(r, 2).Value)) & "_" & _
                     Trim(SafeCStr(ws.Cells(r, 3).Value))
        If UCase(nomComplet) = UCase(Trim(nom)) Then
            GetFonctionAgent = UCase(Trim(SafeCStr(ws.Cells(r, 4).Value)))
            Exit Function
        End If
    Next r
End Function

'--------------------------------------------------------------------
' GetNomMoisFromNum: convertit numero de mois en nom d'onglet
'--------------------------------------------------------------------
Private Function GetNomMoisFromNum(ByVal moisNum As Long) As String
    Dim mSheets() As String
    mSheets = Split(MONTH_SHEETS, ",")
    If moisNum >= 1 And moisNum <= 12 Then
        GetNomMoisFromNum = mSheets(moisNum - 1)
    Else
        GetNomMoisFromNum = ""
    End If
End Function

'--------------------------------------------------------------------
' IsMonthSheet: verifie si un nom de feuille est un onglet mois
'--------------------------------------------------------------------
Public Function IsMonthSheet(ByVal sheetName As String) As Boolean
    Dim mSheets() As String
    Dim i As Long
    mSheets = Split(MONTH_SHEETS, ",")
    IsMonthSheet = False
    For i = 0 To 11
        If UCase(Trim(mSheets(i))) = UCase(Trim(sheetName)) Then
            IsMonthSheet = True
            Exit Function
        End If
    Next i
End Function
