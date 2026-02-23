Attribute VB_Name = "Module_Config_Personnel"
'====================================================================
' MODULE CONFIG_PERSONNEL - Planning_2026_RUNTIME
' Gestion de la feuille Config_Personnel : donnees agents et quotas
' Remplace les quotas hardcodes de Module_SuiviRH.GetQuotas()
'
' COLONNES Config_Personnel:
'   A=Matricule, B=Nom, C=Prenom, D=Fonction, E=DateEntree,
'   F=DateSortie, G=ContratBase, H=PctTemps, I=RegimeCTR,
'   J=QuotaCA, K=QuotaEL, L=QuotaANC, M=QuotaCSoc,
'   N=QuotaDP, O=QuotaCRP, P=HeuresStdJour
'
' DEPENDANCES:
'   - Aucune dependance externe (module autonome)
'   - Consomme par: Module_SuiviRH, Module_Conges_Engine
'
' INSTALLATION:
'   1. Ouvrir Planning_2026_RUNTIME.xlsm
'   2. Alt+F11 > Menu Fichier > Importer > Module_Config_Personnel.bas
'   3. Lancer: Alt+F8 > InitialiserConfigPersonnel > Executer
'   4. Puis: Alt+F8 > MigrerQuotasDepuisSuiviRH > Executer
'====================================================================

Option Explicit

' ---- CONSTANTES ----
Private Const SHEET_NAME As String = "Config_Personnel"
Private Const FIRST_DATA_ROW As Long = 2   ' Ligne 1 = en-tetes
Private Const COL_MATRICULE As Long = 1     ' A
Private Const COL_NOM As Long = 2           ' B
Private Const COL_PRENOM As Long = 3        ' C
Private Const COL_FONCTION As Long = 4      ' D
Private Const COL_DATE_ENTREE As Long = 5   ' E
Private Const COL_DATE_SORTIE As Long = 6   ' F
Private Const COL_CONTRAT As Long = 7       ' G
Private Const COL_PCT_TEMPS As Long = 8     ' H
Private Const COL_REGIME_CTR As Long = 9    ' I
Private Const COL_QUOTA_CA As Long = 10     ' J
Private Const COL_QUOTA_EL As Long = 11     ' K
Private Const COL_QUOTA_ANC As Long = 12    ' L
Private Const COL_QUOTA_CSOC As Long = 13   ' M
Private Const COL_QUOTA_DP As Long = 14     ' N
Private Const COL_QUOTA_CRP As Long = 15    ' O
Private Const COL_HEURES_STD As Long = 16   ' P
Private Const LAST_COL As Long = 16         ' Derniere colonne

' ---- HEADERS ----
Private Const HEADERS As String = "Matricule,Nom,Prenom,Fonction,DateEntree,DateSortie,ContratBase,PctTemps,RegimeCTR,QuotaCA,QuotaEL,QuotaANC,QuotaCSoc,QuotaDP,QuotaCRP,HeuresStdJour"

' ---- HELPER: conversion safe cellule -> String ----
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
' SUB PUBLIQUE 1: InitialiserConfigPersonnel
' Cree la feuille Config_Personnel avec les en-tetes et le formatage
' Si la feuille existe deja, ne fait rien (securite)
'====================================================================
Public Sub InitialiserConfigPersonnel()
    Dim ws As Worksheet
    Dim headerArr() As String
    Dim col As Long
    Dim rng As Range

    ' Verifier si la feuille existe deja
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If Not ws Is Nothing Then
        MsgBox "La feuille '" & SHEET_NAME & "' existe deja." & vbCrLf & _
               "Supprimez-la manuellement avant de reinitialiser.", _
               vbExclamation, "Config_Personnel"
        Exit Sub
    End If

    ' Creer la feuille
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SHEET_NAME

    ' Ecrire les en-tetes
    headerArr = Split(HEADERS, ",")
    For col = 0 To UBound(headerArr)
        ws.Cells(1, col + 1).Value = headerArr(col)
    Next col

    ' Formatage en-tete
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, LAST_COL))
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Largeurs de colonnes
    ws.Columns(COL_MATRICULE).ColumnWidth = 12
    ws.Columns(COL_NOM).ColumnWidth = 22
    ws.Columns(COL_PRENOM).ColumnWidth = 16
    ws.Columns(COL_FONCTION).ColumnWidth = 10
    ws.Columns(COL_DATE_ENTREE).ColumnWidth = 12
    ws.Columns(COL_DATE_SORTIE).ColumnWidth = 12
    ws.Columns(COL_CONTRAT).ColumnWidth = 10
    ws.Columns(COL_PCT_TEMPS).ColumnWidth = 10
    ws.Columns(COL_REGIME_CTR).ColumnWidth = 12
    ws.Columns(COL_QUOTA_CA).ColumnWidth = 9
    ws.Columns(COL_QUOTA_EL).ColumnWidth = 9
    ws.Columns(COL_QUOTA_ANC).ColumnWidth = 9
    ws.Columns(COL_QUOTA_CSOC).ColumnWidth = 9
    ws.Columns(COL_QUOTA_DP).ColumnWidth = 9
    ws.Columns(COL_QUOTA_CRP).ColumnWidth = 9
    ws.Columns(COL_HEURES_STD).ColumnWidth = 12

    ' Format numerique pour les colonnes quotas et %
    ws.Columns(COL_PCT_TEMPS).NumberFormat = "0.00"
    ws.Columns(COL_QUOTA_CA).NumberFormat = "0"
    ws.Columns(COL_QUOTA_EL).NumberFormat = "0"
    ws.Columns(COL_QUOTA_ANC).NumberFormat = "0"
    ws.Columns(COL_QUOTA_CSOC).NumberFormat = "0"
    ws.Columns(COL_QUOTA_DP).NumberFormat = "0"
    ws.Columns(COL_QUOTA_CRP).NumberFormat = "0.00"
    ws.Columns(COL_HEURES_STD).NumberFormat = "0.00"

    ' Format date
    ws.Columns(COL_DATE_ENTREE).NumberFormat = "dd/mm/yyyy"
    ws.Columns(COL_DATE_SORTIE).NumberFormat = "dd/mm/yyyy"

    ' Freeze panes
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' Hauteur en-tete
    ws.Rows(1).RowHeight = 30

    ws.Activate

    MsgBox "Feuille '" & SHEET_NAME & "' creee avec succes." & vbCrLf & _
           "Lancez 'MigrerQuotasDepuisSuiviRH' pour pre-remplir les agents.", _
           vbInformation, "Config_Personnel"
End Sub

'====================================================================
' SUB PUBLIQUE 2: MigrerQuotasDepuisSuiviRH
' Pre-remplit Config_Personnel avec les 25 agents hardcodes
' de Module_SuiviRH (quotas CA, EL, ANC, CSOC, DP, CRP)
' Genere aussi un matricule auto-incremente (CP001, CP002...)
'====================================================================
Public Sub MigrerQuotasDepuisSuiviRH()
    Dim ws As Worksheet
    Dim r As Long
    Dim parts() As String
    Dim nomComplet As String, nom As String, prenom As String
    Dim matricule As String

    ' Verifier que la feuille existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La feuille '" & SHEET_NAME & "' n'existe pas." & vbCrLf & _
               "Lancez d'abord 'InitialiserConfigPersonnel'.", _
               vbCritical, "Config_Personnel"
        Exit Sub
    End If

    ' Verifier si des donnees existent deja
    If Len(Trim(SafeCStr(ws.Cells(FIRST_DATA_ROW, COL_NOM).Value))) > 0 Then
        If MsgBox("Des donnees existent deja dans Config_Personnel." & vbCrLf & _
                  "Voulez-vous AJOUTER les agents manquants (sans doublons) ?", _
                  vbYesNo + vbQuestion, "Config_Personnel") = vbNo Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False

    ' ---- Les 25 agents depuis Module_SuiviRH.GetQuotas() ----
    ' Format: "Nom_Prenom|CA|EL|ANC|CSOC|DP|CRP"
    Dim agentData(0 To 24) As String
    agentData(0) = "Hermann_Claude|24|5|4|2|1|0"
    agentData(1) = "Ben Abdelkader_Yahya|24|3|3|2|0|0"
    agentData(2) = "Bourgeois_Aurore|24|5|4|2|1|8.35"
    agentData(3) = "Ourtioualous_Naaima|24|8|1|2|1|0"
    agentData(4) = "Bozic_Jacqueline|24|5|0|2|0|0"
    agentData(5) = "Youssouf_Roukkiat|24|2|3|2|2|0"
    agentData(6) = "Wielemans_Jennelie|24|5|2|2|0|0"
    agentData(7) = "El Gharbaoui_Sherazade|24|5|3|2|0|0"
    agentData(8) = "Mupika Manga_Caroline|24|4|3|2|2|0"
    agentData(9) = "Ulpat_Victor|24|5|1|2|0|0"
    agentData(10) = "Haouriqui_Mohamed|24|5|1|2|0|0"
    agentData(11) = "Vorst_Julie|24|5|0|2|0|0"
    agentData(12) = "Diallo_Mamadou|24|5|0|2|0|0"
    agentData(13) = "Dela Vega_Edelyn|24|5|1|2|0|7.96"
    agentData(14) = "Ousrout_Salma|24|5|0|2|0|0"
    agentData(15) = "Mutombo Ilunga_Francis|24|1|1|2|0|0"
    agentData(16) = "Bossaert_Marion|24|3|0|2|0|0"
    agentData(17) = "De Bus_Anja|24|5|0|2|0|0"
    agentData(18) = "Adzogble_Charles|24|5|3|2|0|0"
    agentData(19) = "Nana Chamba_Henri|24|2|1|2|0|0"
    agentData(20) = "De Smedt_Sabrina|24|3|0|2|0|0"
    agentData(21) = "AlaHyane_Zahra|24|5|0|2|0|0"
    agentData(22) = "Uwera_Laetitia|24|5|2|2|3|0"
    agentData(23) = "Nayiturikiv_Verene|24|4|1|2|1|0"
    agentData(24) = "Ramack_Sylvie|24|5|7|2|3|10"

    Dim i As Long
    Dim dataFields() As String
    Dim existingNames As Object
    Set existingNames = CreateObject("Scripting.Dictionary")

    ' Collecter les noms existants pour eviter les doublons
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    Dim rr As Long
    For rr = FIRST_DATA_ROW To lastRow
        Dim existingKey As String
        existingKey = UCase(Trim(SafeCStr(ws.Cells(rr, COL_NOM).Value)) & "_" & _
                     Trim(SafeCStr(ws.Cells(rr, COL_PRENOM).Value)))
        If Len(existingKey) > 1 Then
            existingNames.Add existingKey, True
        End If
    Next rr

    ' Trouver la prochaine ligne libre et le prochain numero de matricule
    r = lastRow + 1
    If r < FIRST_DATA_ROW Then r = FIRST_DATA_ROW
    Dim nextMatriculeNum As Long
    nextMatriculeNum = r - FIRST_DATA_ROW + 1

    Dim addedCount As Long
    addedCount = 0

    For i = 0 To 24
        dataFields = Split(agentData(i), "|")
        nomComplet = dataFields(0) ' "Nom_Prenom"

        ' Separer nom et prenom (separateur = "_")
        Dim underscorePos As Long
        underscorePos = InStr(nomComplet, "_")
        If underscorePos > 0 Then
            nom = Left(nomComplet, underscorePos - 1)
            prenom = Mid(nomComplet, underscorePos + 1)
        Else
            nom = nomComplet
            prenom = ""
        End If

        ' Verifier doublon
        Dim checkKey As String
        checkKey = UCase(nom & "_" & prenom)
        If existingNames.Exists(checkKey) Then GoTo NextAgent

        ' Generer matricule
        matricule = "CP" & Format(nextMatriculeNum, "000")
        nextMatriculeNum = nextMatriculeNum + 1

        ' Ecrire la ligne
        ws.Cells(r, COL_MATRICULE).Value = matricule
        ws.Cells(r, COL_NOM).Value = nom
        ws.Cells(r, COL_PRENOM).Value = prenom
        ws.Cells(r, COL_FONCTION).Value = ""             ' A remplir manuellement
        ws.Cells(r, COL_DATE_ENTREE).Value = ""           ' A remplir manuellement
        ws.Cells(r, COL_DATE_SORTIE).Value = ""           ' Vide = actif
        ws.Cells(r, COL_CONTRAT).Value = "CDI"            ' Defaut
        ws.Cells(r, COL_PCT_TEMPS).Value = 1#             ' 100% par defaut
        ws.Cells(r, COL_REGIME_CTR).Value = "NEANT"       ' Defaut
        ws.Cells(r, COL_QUOTA_CA).Value = CDbl(dataFields(1))
        ws.Cells(r, COL_QUOTA_EL).Value = CDbl(dataFields(2))
        ws.Cells(r, COL_QUOTA_ANC).Value = CDbl(dataFields(3))
        ws.Cells(r, COL_QUOTA_CSOC).Value = CDbl(dataFields(4))
        ws.Cells(r, COL_QUOTA_DP).Value = CDbl(dataFields(5))
        ws.Cells(r, COL_QUOTA_CRP).Value = CDbl(dataFields(6))
        ws.Cells(r, COL_HEURES_STD).Value = 7.6           ' Standard hospitalier belge

        existingNames.Add checkKey, True
        addedCount = addedCount + 1
        r = r + 1
NextAgent:
    Next i

    ' Formatage des donnees
    If addedCount > 0 Then
        Dim dataRange As Range
        Set dataRange = ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(r - 1, LAST_COL))
        dataRange.Borders.LineStyle = xlContinuous
        dataRange.Borders.Weight = xlThin
        dataRange.Font.Name = "Calibri"
        dataRange.Font.Size = 11

        ' Alternance couleur de fond pour lisibilite
        Dim rowIdx As Long
        For rowIdx = FIRST_DATA_ROW To r - 1
            If rowIdx Mod 2 = 0 Then
                ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, LAST_COL)).Interior.Color = RGB(234, 241, 251)
            End If
        Next rowIdx
    End If

    Application.ScreenUpdating = True

    MsgBox addedCount & " agents migres avec succes vers Config_Personnel." & vbCrLf & _
           "Completez les champs Fonction, DateEntree et PctTemps manuellement.", _
           vbInformation, "Migration Quotas"
End Sub

'====================================================================
' FONCTION PUBLIQUE 3: GetAgentConfig
' Lit toutes les donnees d'un agent depuis Config_Personnel
' Cherche par NomComplet (format "Nom_Prenom" comme dans le planning)
'
' @param nomAgent  String  Nom au format "Nom_Prenom" (ex: "Hermann_Claude")
' @return Variant  Array(0..15) avec les 16 colonnes, ou Empty si non trouve
'====================================================================
Public Function GetAgentConfig(ByVal nomAgent As String) As Variant
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim nomCheck As String, prenomCheck As String
    Dim clef As String
    Dim result(0 To 15) As Variant

    GetAgentConfig = Empty

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = GetLastDataRow(ws)

    For r = FIRST_DATA_ROW To lastRow
        nomCheck = Trim(SafeCStr(ws.Cells(r, COL_NOM).Value))
        prenomCheck = Trim(SafeCStr(ws.Cells(r, COL_PRENOM).Value))
        clef = nomCheck & "_" & prenomCheck

        If UCase(Trim(nomAgent)) = UCase(clef) Then
            result(0) = SafeCStr(ws.Cells(r, COL_MATRICULE).Value)
            result(1) = nomCheck
            result(2) = prenomCheck
            result(3) = SafeCStr(ws.Cells(r, COL_FONCTION).Value)
            result(4) = ws.Cells(r, COL_DATE_ENTREE).Value
            result(5) = ws.Cells(r, COL_DATE_SORTIE).Value
            result(6) = SafeCStr(ws.Cells(r, COL_CONTRAT).Value)

            If IsNumeric(ws.Cells(r, COL_PCT_TEMPS).Value) Then
                result(7) = CDbl(ws.Cells(r, COL_PCT_TEMPS).Value)
            Else
                result(7) = 1#
            End If

            result(8) = SafeCStr(ws.Cells(r, COL_REGIME_CTR).Value)

            Dim c As Long
            For c = COL_QUOTA_CA To COL_HEURES_STD
                If IsNumeric(ws.Cells(r, c).Value) Then
                    result(c - 1) = CDbl(ws.Cells(r, c).Value)
                Else
                    result(c - 1) = 0#
                End If
            Next c

            GetAgentConfig = result
            Exit Function
        End If
    Next r
End Function

'====================================================================
' FONCTION PUBLIQUE 4: GetAllAgents
' Retourne une Collection de tous les agents actifs
' (DateSortie vide = agent actif)
' Chaque element est un Array(0..15) identique a GetAgentConfig
'
' @return Collection  Collection d'arrays agent
'====================================================================
Public Function GetAllAgents() As Collection
    Dim ws As Worksheet
    Dim col As New Collection
    Dim r As Long, lastRow As Long
    Dim dateSortie As String
    Dim agentArr(0 To 15) As Variant
    Dim c As Long

    Set GetAllAgents = New Collection

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = GetLastDataRow(ws)

    For r = FIRST_DATA_ROW To lastRow
        ' Verifier que la ligne a un nom
        If Len(Trim(SafeCStr(ws.Cells(r, COL_NOM).Value))) = 0 Then GoTo NextRow

        ' Verifier si l'agent est actif (DateSortie vide)
        dateSortie = Trim(SafeCStr(ws.Cells(r, COL_DATE_SORTIE).Value))
        If Len(dateSortie) > 0 Then
            ' DateSortie remplie = agent sorti, on le saute
            GoTo NextRow
        End If

        ' Construire l'array agent
        Dim result(0 To 15) As Variant
        result(0) = SafeCStr(ws.Cells(r, COL_MATRICULE).Value)
        result(1) = Trim(SafeCStr(ws.Cells(r, COL_NOM).Value))
        result(2) = Trim(SafeCStr(ws.Cells(r, COL_PRENOM).Value))
        result(3) = SafeCStr(ws.Cells(r, COL_FONCTION).Value)
        result(4) = ws.Cells(r, COL_DATE_ENTREE).Value
        result(5) = ws.Cells(r, COL_DATE_SORTIE).Value
        result(6) = SafeCStr(ws.Cells(r, COL_CONTRAT).Value)

        If IsNumeric(ws.Cells(r, COL_PCT_TEMPS).Value) Then
            result(7) = CDbl(ws.Cells(r, COL_PCT_TEMPS).Value)
        Else
            result(7) = 1#
        End If

        result(8) = SafeCStr(ws.Cells(r, COL_REGIME_CTR).Value)

        For c = COL_QUOTA_CA To COL_HEURES_STD
            If IsNumeric(ws.Cells(r, c).Value) Then
                result(c - 1) = CDbl(ws.Cells(r, c).Value)
            Else
                result(c - 1) = 0#
            End If
        Next c

        ' Utiliser Nom_Prenom comme cle
        Dim clef As String
        clef = result(1) & "_" & result(2)
        col.Add result, clef
NextRow:
    Next r

    Set GetAllAgents = col
End Function

'====================================================================
' FONCTION PUBLIQUE 5: GetQuotaFromConfig
' Lecture d'un quota specifique pour un agent
'
' @param nomAgent   String  Nom au format "Nom_Prenom"
' @param typeConge  String  Type de conge: "CA","EL","ANC","CSOC","DP","CRP"
' @return Double    Valeur du quota, ou -1 si agent non trouve
'====================================================================
Public Function GetQuotaFromConfig(ByVal nomAgent As String, ByVal typeConge As String) As Double
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim nomCheck As String, prenomCheck As String
    Dim clef As String
    Dim colIdx As Long
    Dim tc As String

    GetQuotaFromConfig = -1#

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    ' Determiner la colonne selon le type de conge
    tc = UCase(Trim(typeConge))
    Select Case tc
        Case "CA": colIdx = COL_QUOTA_CA
        Case "EL": colIdx = COL_QUOTA_EL
        Case "ANC": colIdx = COL_QUOTA_ANC
        Case "CSOC", "C SOC": colIdx = COL_QUOTA_CSOC
        Case "DP": colIdx = COL_QUOTA_DP
        Case "CRP", "CRP1": colIdx = COL_QUOTA_CRP
        Case Else
            GetQuotaFromConfig = 0#
            Exit Function
    End Select

    lastRow = GetLastDataRow(ws)

    For r = FIRST_DATA_ROW To lastRow
        nomCheck = Trim(SafeCStr(ws.Cells(r, COL_NOM).Value))
        prenomCheck = Trim(SafeCStr(ws.Cells(r, COL_PRENOM).Value))
        clef = nomCheck & "_" & prenomCheck

        If UCase(Trim(nomAgent)) = UCase(clef) Then
            If IsNumeric(ws.Cells(r, colIdx).Value) Then
                GetQuotaFromConfig = CDbl(ws.Cells(r, colIdx).Value)
            Else
                GetQuotaFromConfig = 0#
            End If
            Exit Function
        End If
    Next r
End Function

'====================================================================
' FONCTION PUBLIQUE 6: GetHeuresStdFromConfig
' Lecture du HeuresStdJour pour un agent depuis Config_Personnel
'
' @param nomAgent   String  Nom au format "Nom_Prenom"
' @return Double    HeuresStdJour, ou 7.6 si non trouve
'====================================================================
Public Function GetHeuresStdFromConfig(ByVal nomAgent As String) As Double
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim nomCheck As String, prenomCheck As String
    Dim clef As String

    GetHeuresStdFromConfig = 7.6  ' Defaut hospitalier belge

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = GetLastDataRow(ws)

    For r = FIRST_DATA_ROW To lastRow
        nomCheck = Trim(SafeCStr(ws.Cells(r, COL_NOM).Value))
        prenomCheck = Trim(SafeCStr(ws.Cells(r, COL_PRENOM).Value))
        clef = nomCheck & "_" & prenomCheck

        If UCase(Trim(nomAgent)) = UCase(clef) Then
            If IsNumeric(ws.Cells(r, COL_HEURES_STD).Value) And _
               CDbl(ws.Cells(r, COL_HEURES_STD).Value) > 0 Then
                GetHeuresStdFromConfig = CDbl(ws.Cells(r, COL_HEURES_STD).Value)
            End If
            Exit Function
        End If
    Next r
End Function

'====================================================================
' FONCTION PUBLIQUE 7: ValidateConfigPersonnel
' Valide l'integrite des donnees Config_Personnel
' Retourne un message avec les erreurs trouvees, ou "" si tout est OK
'
' @return String  Messages d'erreurs, ou chaine vide si valide
'====================================================================
Public Function ValidateConfigPersonnel() As String
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim erreurs As String
    Dim matricules As Object
    Dim nomsCles As Object
    Dim mat As String, nom As String, prenom As String, clef As String
    Dim pctTemps As Variant
    Dim quotaVal As Variant
    Dim c As Long

    erreurs = ""

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        ValidateConfigPersonnel = "ERREUR: Feuille '" & SHEET_NAME & "' introuvable."
        Exit Function
    End If

    Set matricules = CreateObject("Scripting.Dictionary")
    Set nomsCles = CreateObject("Scripting.Dictionary")
    lastRow = GetLastDataRow(ws)

    For r = FIRST_DATA_ROW To lastRow
        nom = Trim(SafeCStr(ws.Cells(r, COL_NOM).Value))
        prenom = Trim(SafeCStr(ws.Cells(r, COL_PRENOM).Value))
        mat = Trim(SafeCStr(ws.Cells(r, COL_MATRICULE).Value))

        ' Ligne vide = fin
        If Len(nom) = 0 And Len(prenom) = 0 Then GoTo NextValidRow

        ' Verifier matricule non vide
        If Len(mat) = 0 Then
            erreurs = erreurs & "Ligne " & r & ": Matricule vide" & vbCrLf
        End If

        ' Verifier doublon matricule
        If Len(mat) > 0 Then
            If matricules.Exists(mat) Then
                erreurs = erreurs & "Ligne " & r & ": Matricule '" & mat & "' en doublon (deja ligne " & matricules(mat) & ")" & vbCrLf
            Else
                matricules.Add mat, r
            End If
        End If

        ' Verifier doublon Nom_Prenom
        clef = UCase(nom & "_" & prenom)
        If nomsCles.Exists(clef) Then
            erreurs = erreurs & "Ligne " & r & ": Agent '" & nom & "_" & prenom & "' en doublon (deja ligne " & nomsCles(clef) & ")" & vbCrLf
        Else
            nomsCles.Add clef, r
        End If

        ' Verifier PctTemps entre 0 et 1 (ou 0 et 100, tolerance)
        pctTemps = ws.Cells(r, COL_PCT_TEMPS).Value
        If IsNumeric(pctTemps) Then
            If CDbl(pctTemps) < 0 Or CDbl(pctTemps) > 1.5 Then
                ' Peut-etre en pourcent (ex: 75 au lieu de 0.75)
                If CDbl(pctTemps) > 1.5 And CDbl(pctTemps) <= 150 Then
                    erreurs = erreurs & "Ligne " & r & ": PctTemps=" & pctTemps & " (semble etre en %, convertir en decimal: " & Format(CDbl(pctTemps) / 100, "0.00") & ")" & vbCrLf
                Else
                    erreurs = erreurs & "Ligne " & r & ": PctTemps=" & pctTemps & " invalide (doit etre entre 0 et 1)" & vbCrLf
                End If
            End If
        End If

        ' Verifier quotas >= 0
        For c = COL_QUOTA_CA To COL_QUOTA_CRP
            quotaVal = ws.Cells(r, c).Value
            If IsNumeric(quotaVal) Then
                If CDbl(quotaVal) < 0 Then
                    erreurs = erreurs & "Ligne " & r & ": " & ws.Cells(1, c).Value & "=" & quotaVal & " negatif" & vbCrLf
                End If
            End If
        Next c

        ' Verifier HeuresStdJour raisonnable
        Dim hStd As Variant
        hStd = ws.Cells(r, COL_HEURES_STD).Value
        If IsNumeric(hStd) Then
            If CDbl(hStd) <= 0 Or CDbl(hStd) > 12 Then
                erreurs = erreurs & "Ligne " & r & ": HeuresStdJour=" & hStd & " invalide (doit etre entre 0 et 12)" & vbCrLf
            End If
        End If

NextValidRow:
    Next r

    ValidateConfigPersonnel = erreurs
End Function

'====================================================================
' FONCTIONS PRIVEES
'====================================================================

'--------------------------------------------------------------------
' GetLastDataRow: trouve la derniere ligne avec des donnees
'--------------------------------------------------------------------
Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim emptyCount As Long

    GetLastDataRow = FIRST_DATA_ROW - 1
    emptyCount = 0

    For r = FIRST_DATA_ROW To 500
        If Len(Trim(SafeCStr(ws.Cells(r, COL_NOM).Value))) > 0 Or _
           Len(Trim(SafeCStr(ws.Cells(r, COL_MATRICULE).Value))) > 0 Then
            GetLastDataRow = r
            emptyCount = 0
        Else
            emptyCount = emptyCount + 1
            If emptyCount >= 5 Then Exit For
        End If
    Next r
End Function
