Attribute VB_Name = "Module_SyntheseMensuelle"
' Version: v2.0 - 2026-02-23
' Ajouts v2: Colonne "Solde cumule", ligne synthese equipe
' Base: Module_SyntheseMensuelle (Version Finale Definitive)
Option Explicit

'===================================================================================
' MODULE :      Module_SyntheseMensuelle (Version 2.0)
' DESCRIPTION : Calcule les totaux de synthese en se basant EXCLUSIVEMENT
'               sur les informations de la feuille "Config_Codes".
'               v2: Ajoute le solde cumule inter-mois + ligne synthese equipe
'===================================================================================

Private Const MONTH_SHEETS As String = "Janv,Fev,Mars,Avril,Mai,Juin,Juil,Aout,Sept,Oct,Nov,Dec"

Public Sub CalculerSynthesePlannings()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsConfig As Worksheet, wsPlan As Worksheet
    Dim dictCodes As Object, mois As Variant, userChoice As VbMsgBoxResult
    Dim arrOngletsMois As Variant, arrListeMois As Variant

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    arrOngletsMois = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    userChoice = MsgBox("Voulez-vous mettre a jour les soldes pour TOUTE l'annee ?" & vbCrLf & vbCrLf & _
                        "'Oui' = Rapport Annuel (12 mois)" & vbCrLf & _
                        "'Non' = Uniquement pour le mois actif", _
                        vbYesNoCancel + vbQuestion, "Perimetre de la Mise a Jour")

    Select Case userChoice
        Case vbYes: arrListeMois = arrOngletsMois
        Case vbNo:  arrListeMois = Array(ActiveSheet.Name)
        Case vbCancel: GoTo CleanUp
    End Select

    On Error Resume Next
    Set wsConfig = wb.Sheets("Config_Codes")
    On Error GoTo 0
    If wsConfig Is Nothing Then
        MsgBox "ERREUR : La feuille 'Config_Codes' est introuvable.", vbCritical
        GoTo CleanUp
    End If

    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    Dim lastConfigRow As Long: lastConfigRow = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastConfigRow
        Dim code As String: code = Trim(wsConfig.Cells(i, "A").Value)
        If Len(code) > 0 Then
            ' Stocke le Type_Code (Colonne C) et les Heures Digitales (Colonne R)
            dictCodes(code) = Array(wsConfig.Cells(i, "C").Value, wsConfig.Cells(i, "R").Value)
        End If
    Next i

    ' Traiter chaque mois demande
    For Each mois In arrListeMois
        On Error Resume Next
        Set wsPlan = wb.Sheets(mois)
        On Error GoTo 0
        If Not wsPlan Is Nothing Then
            CalculerPourUneFeuille wsPlan, dictCodes
        Else
            Debug.Print "Onglet '" & mois & "' introuvable. Ignore."
        End If
        Set wsPlan = Nothing
    Next mois

    MsgBox "Synthese des plannings terminee !", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub CalculerPourUneFeuille(ByVal ws As Worksheet, ByVal dictCodes As Object)
    Dim colHeuresPrestees As Long, colHeuresAPrester As Long, colSoldeMois As Long
    Dim colSoldeCumule As Long
    Dim colHeuresRecup As Long, colJoursMaladie As Long, colJoursConge As Long, colJoursAbsence As Long

    colHeuresPrestees = TrouverColonne(ws, "Heures prestees")
    If colHeuresPrestees = 0 Then colHeuresPrestees = TrouverColonne(ws, "Heures prest")
    colHeuresAPrester = TrouverColonne(ws, "Heures a prester")
    If colHeuresAPrester = 0 Then colHeuresAPrester = TrouverColonne(ws, "Heures")
    colSoldeMois = TrouverColonne(ws, "Solde du mois")
    colHeuresRecup = TrouverColonne(ws, "Heures a recuperer")
    colJoursMaladie = TrouverColonne(ws, "Jours maladie")
    colJoursConge = TrouverColonne(ws, "Jours conge")
    colJoursAbsence = TrouverColonne(ws, "Jours d'absence")
    If colJoursAbsence = 0 Then colJoursAbsence = TrouverColonne(ws, "Jours absence")

    If colHeuresPrestees = 0 Or colHeuresAPrester = 0 Or colSoldeMois = 0 Then
        MsgBox "ERREUR sur l'onglet '" & ws.Name & "':" & vbCrLf & "Colonnes de synthese introuvables.", vbCritical
        Exit Sub
    End If

    ' --- Trouver ou creer la colonne "Solde cumule" ---
    colSoldeCumule = TrouverColonne(ws, "Solde cumule")
    If colSoldeCumule = 0 Then
        ' Inserer juste apres "Solde du mois"
        colSoldeCumule = colSoldeMois + 1
        ' Decaler les colonnes existantes a droite si necessaire
        ' On cree l'en-tete dans la meme ligne que "Solde du mois"
        Dim headerRow As Long
        headerRow = TrouverLigneHeader(ws, "Solde du mois")
        If headerRow = 0 Then headerRow = 4

        ' Verifier si la colonne suivante est deja occupee par un autre header
        Dim nextColHeader As String
        nextColHeader = Trim(CStr(ws.Cells(headerRow, colSoldeCumule).Value))
        If Len(nextColHeader) > 0 And nextColHeader <> "Solde cumule" Then
            ' Inserer une colonne pour ne pas ecraser
            ws.Columns(colSoldeCumule).Insert Shift:=xlToRight
        End If

        ws.Cells(headerRow, colSoldeCumule).Value = "Solde cumule"
        ws.Cells(headerRow, colSoldeCumule).Font.Bold = True
        ws.Cells(headerRow, colSoldeCumule).Interior.Color = RGB(255, 230, 153) ' Jaune dore
        ws.Cells(headerRow, colSoldeCumule).HorizontalAlignment = xlCenter
        ws.Columns(colSoldeCumule).ColumnWidth = 11
    End If

    Dim lastRow As Long, lastCol As Long, startRow As Long
    Dim arrData As Variant, arrResultats As Variant
    Dim r As Long, c As Long
    Dim codeJour As String, typeCode As String, info As Variant
    Dim heuresJour As Double
    Dim hPrestees As Double, hRecup As Double, jMaladie As Long, jConge As Long, jAbsence As Long

    startRow = 6
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < startRow Then Exit Sub
    lastCol = ws.Cells(startRow - 2, ws.Columns.Count).End(xlToLeft).Column

    arrData = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol)).Value
    ReDim arrResultats(1 To UBound(arrData, 1), 1 To 5)

    For r = 1 To UBound(arrData, 1)
        '--- Reinitialisation des compteurs pour CHAQUE agent ---
        hPrestees = 0: hRecup = 0: jMaladie = 0: jConge = 0: jAbsence = 0

        For c = 4 To UBound(arrData, 2)
            codeJour = Trim(CStr(arrData(r, c)))

            If dictCodes.Exists(codeJour) Then
                info = dictCodes(codeJour)
                typeCode = info(0)
                heuresJour = Val(info(1))

                hPrestees = hPrestees + heuresJour

                Select Case typeCode
                    Case "Recup":         hRecup = hRecup + 1
                    Case "Maladie":       jMaladie = jMaladie + 1
                    Case "Conge":         jConge = jConge + 1
                    Case "SansSolde", "Externe", "Famille", "Exceptionnel": jAbsence = jAbsence + 1
                End Select
            End If
        Next c

        arrResultats(r, 1) = hPrestees: arrResultats(r, 2) = hRecup: arrResultats(r, 3) = jMaladie
        arrResultats(r, 4) = jConge: arrResultats(r, 5) = jAbsence
    Next r

    ' Ecrire les resultats de synthese
    ws.Cells(startRow, colHeuresPrestees).Resize(UBound(arrResultats, 1), 1).Value = Application.Index(arrResultats, 0, 1)
    If colHeuresRecup > 0 Then _
        ws.Cells(startRow, colHeuresRecup).Resize(UBound(arrResultats, 1), 1).Value = Application.Index(arrResultats, 0, 2)
    If colJoursMaladie > 0 Then _
        ws.Cells(startRow, colJoursMaladie).Resize(UBound(arrResultats, 1), 1).Value = Application.Index(arrResultats, 0, 3)
    If colJoursConge > 0 Then _
        ws.Cells(startRow, colJoursConge).Resize(UBound(arrResultats, 1), 1).Value = Application.Index(arrResultats, 0, 4)
    If colJoursAbsence > 0 Then _
        ws.Cells(startRow, colJoursAbsence).Resize(UBound(arrResultats, 1), 1).Value = Application.Index(arrResultats, 0, 5)

    ' Ecrire la formule Solde du mois = Heures prestees - Heures a prester
    ws.Range(ws.Cells(startRow, colSoldeMois), ws.Cells(lastRow, colSoldeMois)).FormulaR1C1 = _
        "=RC" & colHeuresPrestees & "-RC" & colHeuresAPrester

    ' --- SOLDE CUMULE ---
    ' Determiner l'index du mois courant (0-based)
    Dim moisIndex As Long
    moisIndex = GetMoisIndex(ws.Name)

    ' Recuperer le solde cumule du mois precedent (si existe)
    Dim prevSheetName As String
    Dim wsPrev As Worksheet
    Dim colSoldeCumulePrev As Long
    Dim hasPrevCumul As Boolean
    hasPrevCumul = False

    If moisIndex > 0 Then
        prevSheetName = GetMoisName(moisIndex - 1)
        On Error Resume Next
        Set wsPrev = ThisWorkbook.Sheets(prevSheetName)
        On Error GoTo 0
        If Not wsPrev Is Nothing Then
            colSoldeCumulePrev = TrouverColonne(wsPrev, "Solde cumule")
            If colSoldeCumulePrev > 0 Then
                hasPrevCumul = True
            End If
        End If
    End If

    ' Ecrire le solde cumule pour chaque agent
    Dim agentRow As Long
    Dim soldeMois As Double
    Dim soldeCumulePrev As Double
    Dim soldeCumuleVal As Double

    For agentRow = startRow To lastRow
        ' Le solde du mois est dans colSoldeMois
        soldeMois = 0
        On Error Resume Next
        soldeMois = CDbl(ws.Cells(agentRow, colSoldeMois).Value)
        On Error GoTo 0

        If moisIndex = 0 Then
            ' Janvier: solde cumule = solde du mois
            soldeCumuleVal = soldeMois
        Else
            ' Fev-Dec: solde cumule = solde cumule mois precedent + solde mois courant
            soldeCumulePrev = 0
            If hasPrevCumul Then
                ' Chercher la meme ligne dans le mois precedent
                ' Les agents sont aux memes lignes d'un mois a l'autre
                On Error Resume Next
                soldeCumulePrev = CDbl(wsPrev.Cells(agentRow, colSoldeCumulePrev).Value)
                On Error GoTo 0
            End If
            soldeCumuleVal = soldeCumulePrev + soldeMois
        End If

        ws.Cells(agentRow, colSoldeCumule).Value = Round(soldeCumuleVal, 2)
        ws.Cells(agentRow, colSoldeCumule).NumberFormat = "0.00"

        ' Coloration du solde cumule
        If soldeCumuleVal < -5 Then
            ws.Cells(agentRow, colSoldeCumule).Font.Color = RGB(204, 0, 0) ' Rouge
            ws.Cells(agentRow, colSoldeCumule).Font.Bold = True
        ElseIf soldeCumuleVal > 5 Then
            ws.Cells(agentRow, colSoldeCumule).Font.Color = RGB(0, 128, 0) ' Vert
            ws.Cells(agentRow, colSoldeCumule).Font.Bold = True
        Else
            ws.Cells(agentRow, colSoldeCumule).Font.Color = RGB(0, 0, 0)
            ws.Cells(agentRow, colSoldeCumule).Font.Bold = False
        End If
    Next agentRow

    ' --- LIGNE SYNTHESE EQUIPE ---
    EcrireSyntheseEquipe ws, startRow, lastRow, colHeuresAPrester, colHeuresPrestees, colSoldeMois, colSoldeCumule
End Sub


'===================================================================================
' SYNTHESE EQUIPE: Ligne totaux en bas du tableau
'===================================================================================
Private Sub EcrireSyntheseEquipe(ByVal ws As Worksheet, _
                                  ByVal startRow As Long, _
                                  ByVal lastRow As Long, _
                                  ByVal colHAPrester As Long, _
                                  ByVal colHPrestees As Long, _
                                  ByVal colSoldeMois As Long, _
                                  ByVal colSoldeCumule As Long)
    Dim synthRow As Long
    Dim totalHAPrester As Double, totalHPrestees As Double
    Dim totalSoldeMois As Double, totalSoldeCumule As Double
    Dim r As Long

    ' Trouver la premiere ligne vide apres les agents
    synthRow = lastRow + 2

    ' Calculer les totaux
    totalHAPrester = 0
    totalHPrestees = 0
    totalSoldeMois = 0
    totalSoldeCumule = 0

    For r = startRow To lastRow
        Dim nomAgent As String
        nomAgent = Trim(CStr(ws.Cells(r, 2).Value))
        If Len(nomAgent) = 0 Then nomAgent = Trim(CStr(ws.Cells(r, 1).Value))
        If Len(nomAgent) = 0 Then GoTo NextSynthRow

        On Error Resume Next
        totalHAPrester = totalHAPrester + CDbl(ws.Cells(r, colHAPrester).Value)
        totalHPrestees = totalHPrestees + CDbl(ws.Cells(r, colHPrestees).Value)
        totalSoldeMois = totalSoldeMois + CDbl(ws.Cells(r, colSoldeMois).Value)
        totalSoldeCumule = totalSoldeCumule + CDbl(ws.Cells(r, colSoldeCumule).Value)
        On Error GoTo 0
NextSynthRow:
    Next r

    ' Ecrire la ligne synthese
    ws.Cells(synthRow, 1).Value = "TOTAL EQUIPE"
    ws.Cells(synthRow, 1).Font.Bold = True
    ws.Cells(synthRow, 1).Font.Size = 11
    ws.Cells(synthRow, 1).Interior.Color = RGB(31, 78, 121)
    ws.Cells(synthRow, 1).Font.Color = RGB(255, 255, 255)

    ' Etendre le fond bleu sur les colonnes vides entre A et les colonnes de synthese
    Dim colStart As Long
    For colStart = 2 To colHAPrester - 1
        ws.Cells(synthRow, colStart).Interior.Color = RGB(31, 78, 121)
    Next colStart

    ' Total Heures a prester
    ws.Cells(synthRow, colHAPrester).Value = Round(totalHAPrester, 2)
    ws.Cells(synthRow, colHAPrester).NumberFormat = "0.00"
    ws.Cells(synthRow, colHAPrester).Font.Bold = True
    ws.Cells(synthRow, colHAPrester).Interior.Color = RGB(198, 224, 180) ' Vert clair

    ' Total Heures prestees
    ws.Cells(synthRow, colHPrestees).Value = Round(totalHPrestees, 2)
    ws.Cells(synthRow, colHPrestees).NumberFormat = "0.00"
    ws.Cells(synthRow, colHPrestees).Font.Bold = True
    ws.Cells(synthRow, colHPrestees).Interior.Color = RGB(180, 198, 231) ' Bleu clair

    ' Total Solde du mois
    ws.Cells(synthRow, colSoldeMois).Value = Round(totalSoldeMois, 2)
    ws.Cells(synthRow, colSoldeMois).NumberFormat = "0.00"
    ws.Cells(synthRow, colSoldeMois).Font.Bold = True
    If totalSoldeMois < 0 Then
        ws.Cells(synthRow, colSoldeMois).Font.Color = RGB(204, 0, 0)
    Else
        ws.Cells(synthRow, colSoldeMois).Font.Color = RGB(0, 128, 0)
    End If
    ws.Cells(synthRow, colSoldeMois).Interior.Color = RGB(255, 242, 204) ' Jaune pale

    ' Total Solde cumule
    ws.Cells(synthRow, colSoldeCumule).Value = Round(totalSoldeCumule, 2)
    ws.Cells(synthRow, colSoldeCumule).NumberFormat = "0.00"
    ws.Cells(synthRow, colSoldeCumule).Font.Bold = True
    If totalSoldeCumule < -5 Then
        ws.Cells(synthRow, colSoldeCumule).Font.Color = RGB(204, 0, 0)
    ElseIf totalSoldeCumule > 5 Then
        ws.Cells(synthRow, colSoldeCumule).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(synthRow, colSoldeCumule).Font.Color = RGB(0, 0, 0)
    End If
    ws.Cells(synthRow, colSoldeCumule).Interior.Color = RGB(255, 230, 153) ' Jaune dore

    ' Bordure sur la ligne synthese
    Dim rngSynth As Range
    Set rngSynth = ws.Range(ws.Cells(synthRow, 1), ws.Cells(synthRow, colSoldeCumule))
    rngSynth.Borders(xlEdgeTop).LineStyle = xlContinuous
    rngSynth.Borders(xlEdgeTop).Weight = xlMedium
    rngSynth.Borders(xlEdgeBottom).LineStyle = xlDouble
End Sub


'===================================================================================
' HELPERS
'===================================================================================

Private Function TrouverColonne(ws As Worksheet, nomHeader As String) As Long
    Dim searchRange As Range, foundCell As Range
    Set searchRange = ws.Range("A1:AZ5")
    Set foundCell = searchRange.Find(What:=nomHeader, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not foundCell Is Nothing Then TrouverColonne = foundCell.Column Else TrouverColonne = 0
End Function

Private Function TrouverLigneHeader(ws As Worksheet, nomHeader As String) As Long
    Dim searchRange As Range, foundCell As Range
    Set searchRange = ws.Range("A1:AZ5")
    Set foundCell = searchRange.Find(What:=nomHeader, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not foundCell Is Nothing Then TrouverLigneHeader = foundCell.Row Else TrouverLigneHeader = 0
End Function

Private Function GetMoisIndex(ByVal nomMois As String) As Long
    ' Retourne l'index 0-based du mois (0=Janv, 11=Dec)
    Dim mSheets() As String
    Dim i As Long
    mSheets = Split(MONTH_SHEETS, ",")
    For i = 0 To 11
        If UCase(Trim(mSheets(i))) = UCase(Trim(nomMois)) Then
            GetMoisIndex = i
            Exit Function
        End If
    Next i
    GetMoisIndex = -1
End Function

Private Function GetMoisName(ByVal idx As Long) As String
    ' Retourne le nom du mois pour un index 0-based
    Dim mSheets() As String
    If idx < 0 Or idx > 11 Then
        GetMoisName = ""
        Exit Function
    End If
    mSheets = Split(MONTH_SHEETS, ",")
    GetMoisName = mSheets(idx)
End Function
