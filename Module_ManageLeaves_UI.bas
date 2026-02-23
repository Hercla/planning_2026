Attribute VB_Name = "Module_ManageLeaves_UI"
'====================================================================
' MODULE MANAGE LEAVES UI - Planning_2026_RUNTIME
' Interface de gestion des conges : Dashboard Excel + logique
'
' APPROCHE: Feuille "Dashboard_Conges" avec zones de saisie,
'           dropdowns (Data Validation), boutons macro et affichage soldes.
'           Pas besoin de VBProject access (compatible toutes configs).
'
' DEPENDANCES:
'   - Module_Conges_Engine (optionnel, pour soldes/validation/historique)
'   - Module_Config_Personnel (optionnel, pour liste agents)
'   - Module_HeuresTravaillees (optionnel, pour jours ouvrables)
'   - Feuille Personnel (liste agents)
'
' USAGE:
'   Alt+F8 > ShowManageLeavesForm > Executer
'   Ou bouton dans le menu/ruban
'====================================================================

Option Explicit

Private Const DASH_SHEET As String = "Dashboard_Conges"
Private Const MONTH_SHEETS As String = "Janv,Fev,Mars,Avril,Mai,Juin,Juil,Aout,Sept,Oct,Nov,Dec"
Private Const FIRST_EMP_ROW As Long = 5

' ============================================================
' POINT D'ENTREE PRINCIPAL
' ============================================================
Public Sub ShowManageLeavesForm()
    Dim wsDash As Worksheet
    Set wsDash = GetOrCreateDashboard()
    wsDash.Activate
    wsDash.Range("C4").Select
    MsgBox "Dashboard Conges pret !" & vbCrLf & vbCrLf & _
           "1. Selectionnez un agent (C4)" & vbCrLf & _
           "2. Selectionnez le type de conge (C6)" & vbCrLf & _
           "3. Entrez les dates (C8 et C10)" & vbCrLf & _
           "4. Cliquez sur 'Poser le conge'", vbInformation, "Gestion Conges"
End Sub

' ============================================================
' CREATION DU DASHBOARD
' ============================================================
Private Function GetOrCreateDashboard() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DASH_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = DASH_SHEET
        SetupDashboardLayout ws
    End If

    ' Toujours rafraichir la liste des agents
    RefreshAgentDropdown ws

    Set GetOrCreateDashboard = ws
End Function

Private Sub SetupDashboardLayout(ws As Worksheet)
    ' ---- TITRE ----
    ws.Range("B2").value = "GESTION DES CONGES"
    With ws.Range("B2:G2")
        .Merge
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    ' ---- BLOC SAISIE (gauche) ----
    ' Labels
    ws.Range("B4").value = "Agent :"
    ws.Range("B6").value = "Type conge :"
    ws.Range("B8").value = "Date debut :"
    ws.Range("B10").value = "Date fin :"
    ws.Range("B12").value = "Nb jours :"

    ' Formatage labels
    Dim lbl As Variant
    For Each lbl In Array("B4", "B6", "B8", "B10", "B12")
        ws.Range(lbl).Font.Bold = True
        ws.Range(lbl).Font.Size = 11
    Next

    ' Zones de saisie (C4, C6, C8, C10, C12)
    Dim inp As Variant
    For Each inp In Array("C4", "C6", "C8", "C10")
        With ws.Range(inp)
            .Interior.Color = RGB(255, 255, 230)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
            .Borders.Color = RGB(31, 78, 121)
            .Font.Size = 11
            .ColumnWidth = 28
        End With
    Next

    ' C12 = calcul auto (lecture seule)
    With ws.Range("C12")
        .Interior.Color = RGB(230, 240, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 11
        .Font.Bold = True
        .value = "( automatique )"
    End With

    ' ---- DROPDOWN TYPE CONGE ----
    With ws.Range("C6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="CA,EL,ANC,C SOC,DP,CRP,CTR,RCT,MAL,MAT"
    End With

    ' ---- FORMAT DATES ----
    ws.Range("C8").NumberFormat = "dd/mm/yyyy"
    ws.Range("C10").NumberFormat = "dd/mm/yyyy"

    ' ---- BLOC SOLDES (droite) ----
    ws.Range("E4").value = "SOLDES ACTUELS"
    With ws.Range("E4:G4")
        .Merge
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    ' Headers soldes
    ws.Range("E5").value = "Type"
    ws.Range("F5").value = "Acquis"
    ws.Range("G5").value = "Solde"
    ws.Range("E5:G5").Font.Bold = True
    ws.Range("E5:G5").Interior.Color = RGB(200, 220, 240)

    ' Types de conges (E6:E11)
    Dim types As Variant, idx As Long
    types = Array("CA", "EL", "ANC", "C SOC", "DP", "CRP")
    For idx = 0 To 5
        ws.Cells(6 + idx, 5).value = types(idx) ' Col E
        ws.Cells(6 + idx, 6).value = "-"         ' Col F = Acquis
        ws.Cells(6 + idx, 7).value = "-"         ' Col G = Solde
    Next

    ' Bordures bloc soldes
    With ws.Range("E5:G11")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With

    ' ---- BLOC RESULTAT ----
    ws.Range("B14").value = "Solde apres deduction :"
    ws.Range("B14").Font.Bold = True
    With ws.Range("C14")
        .Interior.Color = RGB(230, 240, 255)
        .Font.Size = 14
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .value = "-"
    End With

    ' ---- BOUTONS (via shapes rectangles) ----
    ' Bouton "Poser le conge"
    Dim btnPoser As Shape
    Set btnPoser = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("B16").Left, ws.Range("B16").Top, 160, 35)
    With btnPoser
        .Name = "btnPoserConge"
        .Fill.ForeColor.RGB = RGB(0, 128, 0)
        .TextFrame2.TextRange.Text = "Poser le conge"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "PoserCongeDepuisDashboard"
    End With

    ' Bouton "Recalculer soldes"
    Dim btnRecalc As Shape
    Set btnRecalc = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("D16").Left, ws.Range("D16").Top, 160, 35)
    With btnRecalc
        .Name = "btnRecalculer"
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .TextFrame2.TextRange.Text = "Recalculer soldes"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "RecalculerDepuisDashboard"
    End With

    ' Bouton "Rafraichir soldes agent"
    Dim btnRefresh As Shape
    Set btnRefresh = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("F16").Left, ws.Range("F16").Top, 160, 35)
    With btnRefresh
        .Name = "btnRefreshSoldes"
        .Fill.ForeColor.RGB = RGB(100, 100, 100)
        .TextFrame2.TextRange.Text = "Voir soldes agent"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "RefreshSoldesAgent"
    End With

    ' ---- BLOC HISTORIQUE (bas) ----
    ws.Range("B19").value = "HISTORIQUE RECENT"
    With ws.Range("B19:G19")
        .Merge
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    ' Headers historique
    Dim hHist As Variant, hIdx As Long
    hHist = Array("Date", "Agent", "Type", "Action", "Nb Jours", "Solde Apres")
    For hIdx = 0 To 5
        ws.Cells(20, 2 + hIdx).value = hHist(hIdx)
    Next
    ws.Range("B20:G20").Font.Bold = True
    ws.Range("B20:G20").Interior.Color = RGB(200, 220, 240)
    ws.Range("B20:G30").Borders.LineStyle = xlContinuous
    ws.Range("B20:G30").Borders.Weight = xlThin

    ' Largeurs colonnes
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
End Sub

' ============================================================
' RAFRAICHIR LA LISTE DES AGENTS (dropdown C4)
' ============================================================
Private Sub RefreshAgentDropdown(ws As Worksheet)
    Dim agentList As String
    agentList = BuildAgentListString()

    If agentList <> "" Then
        With ws.Range("C4").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:=agentList
        End With
    End If
End Sub

Private Function BuildAgentListString() As String
    ' Essai 1: Config_Personnel
    Dim wsCP As Worksheet
    On Error Resume Next
    Set wsCP = ThisWorkbook.Sheets("Config_Personnel")
    On Error GoTo 0

    Dim result As String, lastRow As Long, r As Long

    If Not wsCP Is Nothing Then
        lastRow = wsCP.Cells(wsCP.Rows.Count, 2).End(xlUp).row
        For r = 2 To lastRow
            Dim nomCP As String
            nomCP = Trim("" & wsCP.Cells(r, 2).value) & "_" & Trim("" & wsCP.Cells(r, 3).value)
            If Len(nomCP) > 1 Then
                If result <> "" Then result = result & ","
                result = result & nomCP
            End If
        Next r
        If result <> "" Then
            BuildAgentListString = result
            Exit Function
        End If
    End If

    ' Essai 2: Feuille Personnel
    Dim wsPers As Worksheet
    On Error Resume Next
    Set wsPers = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0

    If Not wsPers Is Nothing Then
        lastRow = wsPers.Cells(wsPers.Rows.Count, 1).End(xlUp).row
        result = ""
        For r = 2 To lastRow
            Dim nomP As String
            nomP = Trim("" & wsPers.Cells(r, 1).value)
            If nomP <> "" Then
                If result <> "" Then result = result & ","
                result = result & nomP
            End If
        Next r
    End If

    ' Essai 3: Scanner le premier mois
    If result = "" Then
        Dim wsJanv As Worksheet
        On Error Resume Next
        Set wsJanv = ThisWorkbook.Sheets("Janv")
        On Error GoTo 0
        If Not wsJanv Is Nothing Then
            For r = FIRST_EMP_ROW To 50
                Dim agN As String
                agN = Trim("" & wsJanv.Cells(r, 1).value)
                If agN <> "" And InStr(agN, "Remplacement") = 0 Then
                    If result <> "" Then result = result & ","
                    result = result & agN
                End If
            Next r
        End If
    End If

    BuildAgentListString = result
End Function

' ============================================================
' BOUTON: RAFRAICHIR SOLDES D'UN AGENT
' ============================================================
Public Sub RefreshSoldesAgent()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DASH_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim agentNom As String
    agentNom = Trim("" & ws.Range("C4").value)
    If agentNom = "" Then
        MsgBox "Veuillez d'abord selectionner un agent (cellule C4).", vbExclamation
        Exit Sub
    End If

    ' Lire soldes depuis Soldes_Conges
    Dim wsSoldes As Worksheet
    On Error Resume Next
    Set wsSoldes = ThisWorkbook.Sheets("Soldes_Conges")
    If wsSoldes Is Nothing Then Set wsSoldes = ThisWorkbook.Sheets("Soldes Conges")
    On Error GoTo 0

    If wsSoldes Is Nothing Then
        ' Essayer Module_Conges_Engine
        On Error Resume Next
        Dim typesConge As Variant, t As Long
        typesConge = Array("ca", "el", "anc", "c_soc", "dp", "crp")
        For t = 0 To 5
            Dim solde As Double
            solde = Module_Conges_Engine.GetSoldeAgent(agentNom, CStr(typesConge(t)))
            ws.Cells(6 + t, 7).value = solde
            Dim quota As Double
            quota = Module_Conges_Engine.GetQuotaAgent(agentNom, CStr(typesConge(t)))
            ws.Cells(6 + t, 6).value = quota
            ' Coloration
            If solde < 0 Then
                ws.Cells(6 + t, 7).Font.Color = RGB(204, 0, 0)
            ElseIf solde <= 3 Then
                ws.Cells(6 + t, 7).Font.Color = RGB(200, 150, 0)
            Else
                ws.Cells(6 + t, 7).Font.Color = RGB(0, 128, 0)
            End If
        Next t
        On Error GoTo 0
        Exit Sub
    End If

    ' Lire depuis feuille Soldes_Conges/Soldes Conges
    Dim lastRowS As Long, rS As Long
    lastRowS = wsSoldes.Cells(wsSoldes.Rows.Count, 1).End(xlUp).row
    For rS = 2 To lastRowS
        Dim agS As String
        agS = Trim("" & wsSoldes.Cells(rS, 1).value)
        If agS = "" Then agS = Trim("" & wsSoldes.Cells(rS, 2).value)
        If UCase(agS) = UCase(agentNom) Then
            ' Trouver les colonnes Quota/Solde pour chaque type
            ' Structure Soldes_Conges: CA_Acquis(C), CA_Pris(D), CA_Solde(E), etc.
            ws.Range("F6").value = wsSoldes.Cells(rS, 3).value  ' CA Acquis
            ws.Range("G6").value = wsSoldes.Cells(rS, 5).value  ' CA Solde
            ws.Range("F7").value = wsSoldes.Cells(rS, 6).value  ' EL Acquis
            ws.Range("G7").value = wsSoldes.Cells(rS, 8).value  ' EL Solde
            ws.Range("F8").value = wsSoldes.Cells(rS, 9).value  ' ANC Acquis
            ws.Range("G8").value = wsSoldes.Cells(rS, 11).value ' ANC Solde
            ws.Range("F9").value = wsSoldes.Cells(rS, 12).value ' CSOC Acquis
            ws.Range("G9").value = wsSoldes.Cells(rS, 14).value ' CSOC Solde
            ws.Range("F10").value = wsSoldes.Cells(rS, 15).value ' DP Acquis
            ws.Range("G10").value = wsSoldes.Cells(rS, 17).value ' DP Solde
            ws.Range("F11").value = wsSoldes.Cells(rS, 18).value ' CRP Acquis
            ws.Range("G11").value = wsSoldes.Cells(rS, 20).value ' CRP Solde

            ' Coloration soldes
            Dim ri As Long
            For ri = 6 To 11
                If IsNumeric(ws.Cells(ri, 7).value) Then
                    If ws.Cells(ri, 7).value < 0 Then
                        ws.Cells(ri, 7).Font.Color = RGB(204, 0, 0)
                        ws.Cells(ri, 7).Font.Bold = True
                    ElseIf ws.Cells(ri, 7).value <= 3 Then
                        ws.Cells(ri, 7).Font.Color = RGB(200, 150, 0)
                    Else
                        ws.Cells(ri, 7).Font.Color = RGB(0, 128, 0)
                    End If
                End If
            Next ri
            Exit For
        End If
    Next rS

    ' Rafraichir historique
    RefreshHistorique ws, agentNom
End Sub

' ============================================================
' RAFRAICHIR HISTORIQUE
' ============================================================
Private Sub RefreshHistorique(wsDash As Worksheet, agentNom As String)
    ' Effacer zone historique
    wsDash.Range("B21:G30").ClearContents

    Dim wsHist As Worksheet
    On Error Resume Next
    Set wsHist = ThisWorkbook.Sheets("Historique_Conges")
    On Error GoTo 0
    If wsHist Is Nothing Then Exit Sub

    Dim lastRowH As Long, rH As Long, writeRow As Long
    lastRowH = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).row
    writeRow = 21

    ' Lire les 10 dernieres entrees de cet agent (du bas vers le haut)
    Dim count As Long: count = 0
    For rH = lastRowH To 2 Step -1
        If count >= 10 Then Exit For
        Dim histAgent As String
        histAgent = Trim("" & wsHist.Cells(rH, 4).value) ' Col D = NomComplet
        If histAgent = "" Then histAgent = Trim("" & wsHist.Cells(rH, 3).value)
        If UCase(histAgent) = UCase(agentNom) Then
            wsDash.Cells(writeRow, 2).value = wsHist.Cells(rH, 2).value ' DateHeure
            wsDash.Cells(writeRow, 3).value = histAgent                  ' Agent
            wsDash.Cells(writeRow, 4).value = wsHist.Cells(rH, 5).value ' TypeConge
            wsDash.Cells(writeRow, 5).value = wsHist.Cells(rH, 6).value ' Action
            wsDash.Cells(writeRow, 6).value = wsHist.Cells(rH, 9).value ' NbJours
            wsDash.Cells(writeRow, 7).value = wsHist.Cells(rH, 11).value ' SoldeApres
            writeRow = writeRow + 1
            count = count + 1
        End If
    Next rH
End Sub

' ============================================================
' BOUTON: POSER UN CONGE
' ============================================================
Public Sub PoserCongeDepuisDashboard()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DASH_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Dashboard introuvable. Lancez d'abord ShowManageLeavesForm.", vbCritical
        Exit Sub
    End If

    ' ---- Lire les valeurs ----
    Dim agentNom As String, typeConge As String
    Dim dateDeb As Date, dateFin As Date

    agentNom = Trim("" & ws.Range("C4").value)
    typeConge = Trim("" & ws.Range("C6").value)

    If agentNom = "" Then
        MsgBox "Veuillez selectionner un agent.", vbExclamation
        Exit Sub
    End If
    If typeConge = "" Then
        MsgBox "Veuillez selectionner un type de conge.", vbExclamation
        Exit Sub
    End If

    ' Validation dates
    If Not IsDate(ws.Range("C8").value) Then
        MsgBox "Date de debut invalide. Format attendu: dd/mm/yyyy", vbExclamation
        Exit Sub
    End If
    If Not IsDate(ws.Range("C10").value) Then
        MsgBox "Date de fin invalide. Format attendu: dd/mm/yyyy", vbExclamation
        Exit Sub
    End If

    dateDeb = CDate(ws.Range("C8").value)
    dateFin = CDate(ws.Range("C10").value)

    If dateFin < dateDeb Then
        MsgBox "La date de fin doit etre >= date de debut.", vbExclamation
        Exit Sub
    End If

    ' ---- Calcul nb jours ouvrables ----
    Dim nbJours As Long
    nbJours = CalculerJoursOuvrables(dateDeb, dateFin)
    ws.Range("C12").value = nbJours & " jour(s)"

    ' ---- Validation via Module_Conges_Engine (si disponible) ----
    Dim msg As String
    On Error Resume Next
    Dim valide As Boolean
    valide = Module_Conges_Engine.ValiderPriseConge(agentNom, typeConge, dateDeb, dateFin, msg)
    If Err.Number <> 0 Then
        ' Module_Conges_Engine non disponible, validation basique
        Err.Clear
        valide = True
        msg = ""
    End If
    On Error GoTo 0

    If Not valide Then
        MsgBox "Validation echouee :" & vbCrLf & vbCrLf & msg, vbCritical, "Conge refuse"
        Exit Sub
    End If

    ' ---- Confirmation ----
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Poser " & nbJours & " jour(s) de " & typeConge & " pour " & agentNom & " ?" & vbCrLf & _
                     "Du " & Format(dateDeb, "dd/mm/yyyy") & " au " & Format(dateFin, "dd/mm/yyyy"), _
                     vbYesNo + vbQuestion, "Confirmation")
    If confirm <> vbYes Then Exit Sub

    ' ---- Ecriture dans le planning ----
    ' Retourne le nb reel de jours ecrits (exclut WE, 3/4*, etc.)
    Dim nbJoursEffectifs As Long
    nbJoursEffectifs = EcrireCongesDansPlanning(agentNom, typeConge, dateDeb, dateFin)

    If nbJoursEffectifs = 0 Then
        MsgBox "Aucun jour de travail trouve pour " & agentNom & " entre ces dates." & vbCrLf & _
               "(WE, jours off 4/5, ou conge deja pose)", vbExclamation, "Aucun jour ecrit"
        Exit Sub
    End If

    ' Mettre a jour le nb jours affiche avec le reel
    ws.Range("C12").value = nbJoursEffectifs & " jour(s)"

    ' ---- Ecriture historique (si Module_Conges_Engine dispo) ----
    On Error Resume Next
    Module_Conges_Engine.EcrireHistorique "", agentNom, typeConge, "PRISE", _
        dateDeb, dateFin, nbJoursEffectifs, 0, 0, "Dashboard", ""
    On Error GoTo 0

    ' ---- Recalcul soldes ----
    On Error Resume Next
    Module_Conges_Engine.RecalculerTousSoldes
    On Error GoTo 0

    ' ---- Refresh UI ----
    RefreshSoldesAgent

    ' ---- Solde apres ----
    On Error Resume Next
    Dim soldeApres As Double
    Dim typeCode As String
    typeCode = LCase(Replace(typeConge, " ", "_"))
    If typeCode = "c_soc" Then typeCode = "c_soc"
    soldeApres = Module_Conges_Engine.GetSoldeAgent(agentNom, typeCode)
    ws.Range("C14").value = soldeApres & " jour(s)"
    If soldeApres < 0 Then
        ws.Range("C14").Font.Color = RGB(204, 0, 0)
    Else
        ws.Range("C14").Font.Color = RGB(0, 128, 0)
    End If
    On Error GoTo 0

    MsgBox "Conge pose avec succes !" & vbCrLf & _
           agentNom & " : " & typeConge & " x " & nbJoursEffectifs & "j" & vbCrLf & _
           "Du " & Format(dateDeb, "dd/mm/yyyy") & " au " & Format(dateFin, "dd/mm/yyyy"), _
           vbInformation, "Succes"
End Sub

' ============================================================
' ECRIRE CODES CONGES DANS LES FEUILLES PLANNING
' ============================================================
Private Function EcrireCongesDansPlanning(agentNom As String, typeConge As String, _
                                          dateDeb As Date, dateFin As Date) As Long
    ' Retourne le nombre de jours effectivement ecrits
    Dim mSheets() As String
    mSheets = Split(MONTH_SHEETS, ",")

    Dim currentDate As Date
    Dim joursEcrits As Long: joursEcrits = 0
    currentDate = dateDeb

    Application.ScreenUpdating = False
    Do While currentDate <= dateFin
        ' Seulement les jours ouvrables (Lun-Ven)
        If Weekday(currentDate, vbMonday) < 6 Then
            Dim moisIdx As Long
            moisIdx = Month(currentDate) - 1 ' 0-based

            Dim wsPlan As Worksheet
            On Error Resume Next
            Set wsPlan = ThisWorkbook.Sheets(mSheets(moisIdx))
            On Error GoTo 0

            If Not wsPlan Is Nothing Then
                ' Trouver la ligne de l'agent
                Dim agRow As Long
                agRow = FindAgentRow(wsPlan, agentNom)

                If agRow > 0 Then
                    ' Trouver la colonne du jour
                    Dim dayCol As Long
                    dayCol = FindDayColumn(wsPlan, Day(currentDate))

                    If dayCol > 0 Then
                        ' Ecrire le code conge SEULEMENT si la cellule contient
                        ' un vrai horaire de travail (pas WE, 3/4*, conge existant, vide)
                        Dim existingVal As String
                        existingVal = Trim("" & wsPlan.Cells(agRow, dayCol).value)
                        If IsHoraireTravail(existingVal) Then
                            wsPlan.Cells(agRow, dayCol).value = typeConge
                            joursEcrits = joursEcrits + 1
                        End If
                    End If
                End If
            End If
        End If

        currentDate = currentDate + 1
    Loop
    Application.ScreenUpdating = True

    EcrireCongesDansPlanning = joursEcrits
End Function

Private Function FindAgentRow(ws As Worksheet, agentNom As String) As Long
    Dim r As Long
    For r = FIRST_EMP_ROW To 50
        Dim cellVal As String
        cellVal = Trim("" & ws.Cells(r, 1).value)
        If UCase(cellVal) = UCase(agentNom) Then
            FindAgentRow = r
            Exit Function
        End If
    Next r
    FindAgentRow = 0
End Function

Private Function FindDayColumn(ws As Worksheet, dayNum As Long) As Long
    Dim c As Long
    For c = 3 To 40
        If IsNumeric(ws.Cells(4, c).value) Then
            If CLng(ws.Cells(4, c).value) = dayNum Then
                FindDayColumn = c
                Exit Function
            End If
        End If
    Next c
    FindDayColumn = 0
End Function

' ============================================================
' DETECTION HORAIRE DE TRAVAIL
' Retourne True si la cellule contient un vrai shift (ex: "8:30",
' "C 15", "7 15:30", "12 20"). Retourne False si: vide, WE,
' fraction (3/4*, 4/5), code conge existant, RHS, etc.
' ============================================================
Private Function IsHoraireTravail(cellVal As String) As Boolean
    IsHoraireTravail = False

    ' Vide ou zero
    If cellVal = "" Or cellVal = "0" Then Exit Function

    Dim upper As String
    upper = UCase(cellVal)

    ' Weekend
    If upper = "WE" Then Exit Function

    ' Fractions (jour off temps partiel): contient "/"
    If InStr(cellVal, "/") > 0 Then Exit Function

    ' Codes conge/absence connus -> ne pas ecraser
    Dim nonWork As Variant, code As Variant
    nonWork = Array("CA", "EL", "ANC", "DP", "CRP", "CTR", "RCT", _
                    "MAL", "MAT", "RHS", "RV", "FORM", "DET", "RECUP", _
                    "C SOC", "MAL-GAR", "MAL-MUT")
    For Each code In nonWork
        If upper = CStr(code) Then Exit Function
    Next code

    ' Contient ":" -> format horaire (8:30, 7:15 15:45, etc.)
    If InStr(cellVal, ":") > 0 Then
        IsHoraireTravail = True
        Exit Function
    End If

    ' Commence par "C " + chiffre -> code shift (C 15, C 20, C 20 E)
    If Left(upper, 2) = "C " And Len(upper) >= 3 Then
        If Mid(upper, 3, 1) >= "0" And Mid(upper, 3, 1) <= "9" Then
            IsHoraireTravail = True
            Exit Function
        End If
    End If

    ' Commence par un chiffre -> horaire numerique (7 13, 12 20, 8 16:30)
    If Len(cellVal) > 0 Then
        If Left(cellVal, 1) >= "0" And Left(cellVal, 1) <= "9" Then
            IsHoraireTravail = True
            Exit Function
        End If
    End If

    ' Tout le reste: code inconnu -> par securite, ne pas ecraser
End Function

' ============================================================
' BOUTON: RECALCULER TOUS LES SOLDES
' ============================================================
Public Sub RecalculerDepuisDashboard()
    On Error Resume Next

    ' Essai Module_Conges_Engine
    Module_Conges_Engine.RecalculerTousSoldes
    If Err.Number = 0 Then
        MsgBox "Soldes recalcules avec succes (Module_Conges_Engine) !", vbInformation
        On Error GoTo 0
        ' Refresh si un agent est selectionne
        RefreshSoldesAgent
        Exit Sub
    End If
    Err.Clear

    ' Essai Module_SuiviRH
    Module_SuiviRH.GenererSuiviRH
    If Err.Number = 0 Then
        MsgBox "Suivi RH regenere avec succes !", vbInformation
        On Error GoTo 0
        RefreshSoldesAgent
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    MsgBox "Aucun module de recalcul disponible." & vbCrLf & _
           "Verifiez que Module_Conges_Engine ou Module_SuiviRH est importe.", vbExclamation
End Sub

' ============================================================
' CALCUL JOURS OUVRABLES (Lun-Ven, hors feries)
' ============================================================
Private Function CalculerJoursOuvrables(dateDeb As Date, dateFin As Date) As Long
    Dim count As Long: count = 0
    Dim d As Date

    ' Charger feries belges si possible
    Dim feries As Object
    Set feries = Nothing
    On Error Resume Next
    Dim annee As Long
    annee = Year(dateDeb)
    Set feries = Module_Planning_Core.BuildFeriesBE(annee)
    On Error GoTo 0

    For d = dateDeb To dateFin
        If Weekday(d, vbMonday) < 6 Then ' Lun-Ven
            Dim isFerie As Boolean: isFerie = False
            If Not feries Is Nothing Then
                isFerie = feries.Exists(CStr(CLng(d)))
                If Not isFerie Then isFerie = feries.Exists(Format(d, "yyyy-mm-dd"))
                If Not isFerie Then isFerie = feries.Exists(CStr(d))
            End If
            If Not isFerie Then count = count + 1
        End If
    Next d

    CalculerJoursOuvrables = count
End Function
