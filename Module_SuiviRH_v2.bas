Attribute VB_Name = "Module_SuiviRH"
'====================================================================
' MODULE SUIVI RH - CTR 2026 (v2)
' Genere automatiquement les onglets d'analyse RH
' depuis les donnees du Planning RUNTIME
'
' CHANGEMENTS v2:
'   - GetQuotas() lit depuis Config_Personnel avec fallback legacy
'   - GenererSuiviRH() ne supprime plus les onglets
'     -> efface les DONNEES (A2:...) mais garde les en-tetes
'   - Colonne "Derniere MAJ" ajoutee sur chaque onglet
'
' DEPENDANCES:
'   - Module_Config_Personnel (optionnel, pour lecture quotas dynamiques)
'====================================================================

Option Explicit

' ---- CONSTANTES ----
Private Const MONTH_SHEETS As String = "Janv,Fev,Mars,Avril,Mai,Juin,Juil,Aout,Sept,Oct,Nov,Dec"
Private Const MONTH_NAMES As String = "Janvier,Fevrier,Mars,Avril,Mai,Juin,Juillet,Aout,Septembre,Octobre,Novembre,Decembre"
Private Const FIRST_EMP_ROW As Long = 6
Private Const FIRST_DAY_COL As Long = 3  ' Col C

' ---- QUOTAS INDIVIDUELS ----
Private Type AgentQuota
    Nom As String
    ca As Double
    el As Double
    anc As Double
    CSOC As Double
    dp As Double
    CRP1 As Double
End Type

' ---- v2: LECTURE QUOTAS DEPUIS CONFIG_PERSONNEL OU FALLBACK ----
Private Function GetQuotas() As Collection
    ' Essai 1: lire depuis la feuille Config_Personnel
    On Error Resume Next
    Dim wsCP As Worksheet
    Set wsCP = ThisWorkbook.Sheets("Config_Personnel")
    On Error GoTo 0

    If Not wsCP Is Nothing Then
        Set GetQuotas = GetQuotasFromConfig(wsCP)
        If GetQuotas.Count > 0 Then Exit Function
    End If

    ' Fallback: ancien code hardcode
    Set GetQuotas = GetQuotasLegacy()
End Function

Private Function GetQuotasFromConfig(wsCP As Worksheet) As Collection
    Dim col As New Collection
    Dim lastRow As Long, r As Long
    lastRow = wsCP.Cells(wsCP.Rows.Count, 1).End(xlUp).row

    If lastRow < 2 Then
        Set GetQuotasFromConfig = col
        Exit Function
    End If

    Dim q As AgentQuota
    For r = 2 To lastRow
        Dim nomVal As String, prenomVal As String
        nomVal = Trim("" & wsCP.Cells(r, 2).value)    ' Col B = Nom
        prenomVal = Trim("" & wsCP.Cells(r, 3).value)  ' Col C = Prenom

        If nomVal <> "" Then
            q.Nom = nomVal & "_" & prenomVal

            ' Cols J-O = QuotaCA, QuotaEL, QuotaANC, QuotaCSoc, QuotaDP, QuotaCRP
            q.ca = Val("" & wsCP.Cells(r, 10).value)
            q.el = Val("" & wsCP.Cells(r, 11).value)
            q.anc = Val("" & wsCP.Cells(r, 12).value)
            q.CSOC = Val("" & wsCP.Cells(r, 13).value)
            q.dp = Val("" & wsCP.Cells(r, 14).value)
            q.CRP1 = Val("" & wsCP.Cells(r, 15).value)

            On Error Resume Next
            col.Add q, q.Nom
            On Error GoTo 0
        End If
    Next r

    Set GetQuotasFromConfig = col
End Function

Private Function GetQuotasLegacy() As Collection
    Dim col As New Collection

    ' Format: Nom, CA, EL, ANC, CSOC, DP, CRP1
    AddQuota col, "Hermann_Claude", 24, 5, 4, 2, 1, 0
    AddQuota col, "Ben Abdelkader_Yahya", 24, 3, 3, 2, 0, 0
    AddQuota col, "Bourgeois_Aurore", 24, 5, 4, 2, 1, 8.35
    AddQuota col, "Ourtioualous_Naaima", 24, 8, 1, 2, 1, 0
    AddQuota col, "Bozic_Jacqueline", 24, 5, 0, 2, 0, 0
    AddQuota col, "Youssouf_Roukkiat", 24, 2, 3, 2, 2, 0
    AddQuota col, "Wielemans_Jennelie", 24, 5, 2, 2, 0, 0
    AddQuota col, "El Gharbaoui_Sherazade", 24, 5, 3, 2, 0, 0
    AddQuota col, "Mupika Manga_Caroline", 24, 4, 3, 2, 2, 0
    AddQuota col, "Ulpat_Victor", 24, 5, 1, 2, 0, 0
    AddQuota col, "Haouriqui_Mohamed", 24, 5, 1, 2, 0, 0
    AddQuota col, "Vorst_Julie", 24, 5, 0, 2, 0, 0
    AddQuota col, "Diallo_Mamadou", 24, 5, 0, 2, 0, 0
    AddQuota col, "Dela Vega_Edelyn", 24, 5, 1, 2, 0, 7.96
    AddQuota col, "Ousrout_Salma", 24, 5, 0, 2, 0, 0
    AddQuota col, "Mutombo Ilunga_Francis", 24, 1, 1, 2, 0, 0
    AddQuota col, "Bossaert_Marion", 24, 3, 0, 2, 0, 0
    AddQuota col, "De Bus_Anja", 24, 5, 0, 2, 0, 0
    AddQuota col, "Adzogble_Charles", 24, 5, 3, 2, 0, 0
    AddQuota col, "Nana Chamba_Henri", 24, 2, 1, 2, 0, 0
    AddQuota col, "De Smedt_Sabrina", 24, 3, 0, 2, 0, 0
    AddQuota col, "AlaHyane_Zahra", 24, 5, 0, 2, 0, 0
    AddQuota col, "Uwera_Laetitia", 24, 5, 2, 2, 3, 0
    AddQuota col, "Nayiturikiv_Verene", 24, 4, 1, 2, 1, 0
    AddQuota col, "Ramack_Sylvie", 24, 5, 7, 2, 3, 10

    Set GetQuotasLegacy = col
End Function

Private Sub AddQuota(col As Collection, n As String, ca As Double, el As Double, _
    anc As Double, cs As Double, dp As Double, crp As Double)
    Dim q As AgentQuota
    q.Nom = n: q.ca = ca: q.el = el: q.anc = anc
    q.CSOC = cs: q.dp = dp: q.CRP1 = crp
    col.Add q, n
End Sub

' ---- CLASSIFICATION DES CODES ----
Private Function ClassifyCode(code As String) As String
    Dim c As String
    c = Trim(UCase(code))

    If c = "" Or c = "0" Then ClassifyCode = "empty": Exit Function

    Select Case True
        Case c = "CA": ClassifyCode = "ca"
        Case c = "EL": ClassifyCode = "el"
        Case c = "ANC": ClassifyCode = "anc"
        Case c = "C SOC": ClassifyCode = "c_soc"
        Case c = "DP": ClassifyCode = "dp"
        Case c = "CTR": ClassifyCode = "ctr"
        Case c = "RCT": ClassifyCode = "rct"
        Case c = "RV": ClassifyCode = "rv"
        Case Left(c, 3) = "RHS": ClassifyCode = "rhs"
        Case c = "WE": ClassifyCode = "we"
        Case c = "DECES": ClassifyCode = "deces"
        Case Left(c, 7) = "MAL-GAR", Left(c, 7) = "MAL-MUT", Left(c, 3) = "MUT"
            ClassifyCode = "maladie"
        Case Left(c, 7) = "MAT-EMP", Left(c, 7) = "MAT-MUT"
            ClassifyCode = "maladie"
        Case Left(c, 7) = "PAT-EMP", Left(c, 7) = "PAT-MUT"
            ClassifyCode = "maladie"
        Case Left(c, 3) = "CRP": ClassifyCode = "crp"
        Case Left(c, 1) = "F" And InStr(c, "-") > 0: ClassifyCode = "ferie"
        Case Left(c, 1) = "R" And InStr(c, "-") > 0: ClassifyCode = "ferie"
        Case InStr(c, ":") > 0: ClassifyCode = "work"  ' horaires
        Case Left(c, 2) = "C ": ClassifyCode = "work"  ' coupes
        Case Else: ClassifyCode = "other"
    End Select
End Function

Private Function IsWeekend(ws As Worksheet, dayCol As Long) As Boolean
    Dim dayName As String
    dayName = Trim(UCase("" & ws.Cells(3, dayCol).value))
    IsWeekend = (dayName = "SAM" Or dayName = "DIM")
End Function

Private Function IsSaturday(ws As Worksheet, dayCol As Long) As Boolean
    IsSaturday = (Trim(UCase("" & ws.Cells(3, dayCol).value)) = "SAM")
End Function

Private Function IsSunday(ws As Worksheet, dayCol As Long) As Boolean
    IsSunday = (Trim(UCase("" & ws.Cells(3, dayCol).value)) = "DIM")
End Function

' ---- COULEURS ET FORMATAGE ----
Private Sub FormatHeader(ws As Worksheet, lastCol As Long, Optional r As Long = 1)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Private Sub FormatTable(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlThin
    rng.Font.Name = "Calibri"
    rng.Font.Size = 11
    ws.Rows(1).RowHeight = 30
End Sub

' ---- v2: GET OR CREATE SHEET (ne supprime jamais) ----
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Creer la feuille
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    Else
        ' Effacer les donnees mais garder les en-tetes (ligne 1)
        Dim lastR As Long, lastC As Long
        lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastR >= 2 And lastC >= 1 Then
            ws.Range(ws.Cells(2, 1), ws.Cells(lastR, lastC)).ClearContents
            ws.Range(ws.Cells(2, 1), ws.Cells(lastR, lastC)).Interior.Pattern = xlNone
            ws.Range(ws.Cells(2, 1), ws.Cells(lastR, lastC)).Font.Color = RGB(0, 0, 0)
            ws.Range(ws.Cells(2, 1), ws.Cells(lastR, lastC)).Font.Bold = False
        End If
    End If

    Set GetOrCreateSheet = ws
End Function

' ======== MACRO PRINCIPALE ========
Public Sub GenererSuiviRH()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    Dim mSheets() As String, mNames() As String
    mSheets = Split(MONTH_SHEETS, ",")
    mNames = Split(MONTH_NAMES, ",")

    Dim quotas As Collection
    Set quotas = GetQuotas()

    ' ---- Collecter les donnees ----
    Dim agents As New Collection
    Dim agentOrder() As String
    ReDim agentOrder(0)
    Dim agentCount As Long: agentCount = 0

    Dim dictCA As Object, dictEL As Object, dictANC As Object
    Dim dictCSOC As Object, dictDP As Object, dictCRP As Object
    Dim dictMAL As Object, dictWORK As Object
    Dim dictWE_S As Object, dictWE_D As Object

    Set dictCA = CreateObject("Scripting.Dictionary")
    Set dictEL = CreateObject("Scripting.Dictionary")
    Set dictANC = CreateObject("Scripting.Dictionary")
    Set dictCSOC = CreateObject("Scripting.Dictionary")
    Set dictDP = CreateObject("Scripting.Dictionary")
    Set dictCRP = CreateObject("Scripting.Dictionary")
    Set dictMAL = CreateObject("Scripting.Dictionary")
    Set dictWORK = CreateObject("Scripting.Dictionary")
    Set dictWE_S = CreateObject("Scripting.Dictionary")
    Set dictWE_D = CreateObject("Scripting.Dictionary")

    Dim m As Long, r As Long, c As Long
    Dim ws As Worksheet, agName As String, cellVal As String, cType As String
    Dim key As String, numDays As Long

    ' Scan all 12 months
    For m = 0 To 11
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m))
        On Error GoTo ErrHandler
        If ws Is Nothing Then GoTo NextMonth

        numDays = 0
        For c = FIRST_DAY_COL To FIRST_DAY_COL + 31
            If IsNumeric(ws.Cells(4, c).value) And ws.Cells(4, c).value <> "" Then
                numDays = numDays + 1
            End If
        Next c
        If numDays = 0 Then numDays = 31

        For r = FIRST_EMP_ROW To 50
            agName = Trim("" & ws.Cells(r, 1).value)
            If agName = "" Then GoTo NextRow
            If InStr(agName, "Remplacement") > 0 Then GoTo NextRow
            If agName = "Us Nuit" Or agName = "Noel_Elisabeth" Then GoTo NextRow

            If Not dictCA.Exists(agName) Then
                Dim arr(0 To 11) As Double
                dictCA.Add agName, arr
                dictEL.Add agName, arr
                dictANC.Add agName, arr
                dictCSOC.Add agName, arr
                dictDP.Add agName, arr
                dictCRP.Add agName, arr
                dictMAL.Add agName, arr
                dictWORK.Add agName, arr
                dictWE_S.Add agName, arr
                dictWE_D.Add agName, arr
                agentCount = agentCount + 1
                ReDim Preserve agentOrder(agentCount - 1)
                agentOrder(agentCount - 1) = agName
            End If

            For c = FIRST_DAY_COL To FIRST_DAY_COL + numDays - 1
                cellVal = Trim("" & ws.Cells(r, c).value)
                If cellVal = "" Then GoTo NextCol

                cType = ClassifyCode(cellVal)
                Dim tmpArr() As Double

                Select Case cType
                    Case "ca"
                        tmpArr = dictCA(agName): tmpArr(m) = tmpArr(m) + 1: dictCA(agName) = tmpArr
                    Case "el"
                        tmpArr = dictEL(agName): tmpArr(m) = tmpArr(m) + 1: dictEL(agName) = tmpArr
                    Case "anc"
                        tmpArr = dictANC(agName): tmpArr(m) = tmpArr(m) + 1: dictANC(agName) = tmpArr
                    Case "c_soc"
                        tmpArr = dictCSOC(agName): tmpArr(m) = tmpArr(m) + 1: dictCSOC(agName) = tmpArr
                    Case "dp"
                        tmpArr = dictDP(agName): tmpArr(m) = tmpArr(m) + 1: dictDP(agName) = tmpArr
                    Case "crp"
                        tmpArr = dictCRP(agName): tmpArr(m) = tmpArr(m) + 1: dictCRP(agName) = tmpArr
                    Case "maladie"
                        tmpArr = dictMAL(agName): tmpArr(m) = tmpArr(m) + 1: dictMAL(agName) = tmpArr
                    Case "work"
                        tmpArr = dictWORK(agName): tmpArr(m) = tmpArr(m) + 1: dictWORK(agName) = tmpArr
                        If IsSaturday(ws, c) Then
                            tmpArr = dictWE_S(agName): tmpArr(m) = tmpArr(m) + 1: dictWE_S(agName) = tmpArr
                        End If
                        If IsSunday(ws, c) Then
                            tmpArr = dictWE_D(agName): tmpArr(m) = tmpArr(m) + 1: dictWE_D(agName) = tmpArr
                        End If
                End Select
NextCol:
            Next c
NextRow:
        Next r
NextMonth:
        Set ws = Nothing
    Next m

    ' ---- CREER/METTRE A JOUR LES ONGLETS (v2: jamais supprimer) ----

    ' === Sheet 1: Soldes Conges ===
    Dim ws1 As Worksheet
    Set ws1 = GetOrCreateSheet("Soldes Conges")

    Dim headers As Variant
    headers = Array("Agent", "CA Quota", "CA Pris", "CA Reste", _
        "EL Quota", "EL Pris", "EL Reste", "ANC Quota", "ANC Pris", "ANC Reste", _
        "C SOC Quota", "C SOC Pris", "C SOC Reste", "DP Quota", "DP Pris", "DP Reste", _
        "CRP1 Quota", "CRP1 Pris", "CRP1 Reste", "Total Quota", "Total Pris", "Total Reste", _
        "Derniere MAJ")

    For c = 0 To UBound(headers)
        ws1.Cells(1, c + 1).value = headers(c)
    Next c
    FormatHeader ws1, UBound(headers) + 1

    Dim row As Long: row = 2
    Dim i As Long
    For i = 0 To agentCount - 1
        agName = agentOrder(i)
        ws1.Cells(row, 1).value = agName

        Dim q As AgentQuota
        On Error Resume Next
        q = quotas(agName)
        If Err.Number <> 0 Then
            q.ca = 24: q.el = 5: q.anc = 0: q.CSOC = 2: q.dp = 0: q.CRP1 = 0
            Err.Clear
        End If
        On Error GoTo ErrHandler

        Dim caT As Double, elT As Double, ancT As Double
        Dim csT As Double, dpT As Double, crpT As Double
        Dim tmpA() As Double

        tmpA = dictCA(agName): caT = SumArr(tmpA)
        tmpA = dictEL(agName): elT = SumArr(tmpA)
        tmpA = dictANC(agName): ancT = SumArr(tmpA)
        tmpA = dictCSOC(agName): csT = SumArr(tmpA)
        tmpA = dictDP(agName): dpT = SumArr(tmpA)
        tmpA = dictCRP(agName): crpT = SumArr(tmpA)

        ws1.Cells(row, 2) = q.ca: ws1.Cells(row, 3) = caT
        ws1.Cells(row, 4) = q.ca - caT
        ws1.Cells(row, 5) = q.el: ws1.Cells(row, 6) = elT
        ws1.Cells(row, 7) = q.el - elT
        ws1.Cells(row, 8) = q.anc: ws1.Cells(row, 9) = ancT
        ws1.Cells(row, 10) = q.anc - ancT
        ws1.Cells(row, 11) = q.CSOC: ws1.Cells(row, 12) = csT
        ws1.Cells(row, 13) = q.CSOC - csT
        ws1.Cells(row, 14) = q.dp: ws1.Cells(row, 15) = dpT
        ws1.Cells(row, 16) = q.dp - dpT
        ws1.Cells(row, 17) = q.CRP1: ws1.Cells(row, 18) = crpT
        ws1.Cells(row, 19) = q.CRP1 - crpT

        Dim tQ As Double, tT As Double
        tQ = q.ca + q.el + q.anc + q.CSOC + q.dp + q.CRP1
        tT = caT + elT + ancT + csT + dpT + crpT
        ws1.Cells(row, 20) = tQ: ws1.Cells(row, 21) = tT
        ws1.Cells(row, 22) = tQ - tT

        ' v2: Derniere MAJ
        ws1.Cells(row, 23) = Now()
        ws1.Cells(row, 23).NumberFormat = "dd/mm/yyyy hh:mm"

        Dim colIdx As Variant
        For Each colIdx In Array(4, 7, 10, 13, 16, 19, 22)
            If ws1.Cells(row, colIdx).value < 0 Then
                ws1.Cells(row, colIdx).Font.Color = RGB(204, 0, 0)
                ws1.Cells(row, colIdx).Font.Bold = True
                ws1.Cells(row, colIdx).Interior.Color = RGB(255, 199, 206)
            ElseIf ws1.Cells(row, colIdx).value > 0 Then
                ws1.Cells(row, colIdx).Font.Color = RGB(0, 102, 0)
            End If
        Next

        row = row + 1
    Next i

    FormatTable ws1, row - 1, 23
    ws1.Columns("A").ColumnWidth = 28
    ws1.Range("B1:V1").ColumnWidth = 10
    ws1.Columns("W").ColumnWidth = 18
    ws1.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' === Sheet 2: Absenteisme Maladie ===
    Dim ws2 As Worksheet
    Set ws2 = GetOrCreateSheet("Absenteisme")

    Dim h2 As Variant
    h2 = Array("Agent", "Jan", "Fev", "Mar", "Avr", "Mai", "Jun", _
        "Jul", "Aou", "Sep", "Oct", "Nov", "Dec", "T1", "T2", "T3", "T4", "Annee", "Derniere MAJ")
    For c = 0 To UBound(h2)
        ws2.Cells(1, c + 1).value = h2(c)
    Next c
    FormatHeader ws2, UBound(h2) + 1

    row = 2
    For i = 0 To agentCount - 1
        agName = agentOrder(i)
        ws2.Cells(row, 1).value = agName

        tmpA = dictMAL(agName)
        Dim mTotal As Double: mTotal = 0
        For m = 0 To 11
            ws2.Cells(row, m + 2).value = tmpA(m)
            mTotal = mTotal + tmpA(m)
            If tmpA(m) > 5 Then
                ws2.Cells(row, m + 2).Font.Color = RGB(204, 0, 0)
                ws2.Cells(row, m + 2).Font.Bold = True
            End If
        Next m

        ws2.Cells(row, 14) = tmpA(0) + tmpA(1) + tmpA(2)
        ws2.Cells(row, 15) = tmpA(3) + tmpA(4) + tmpA(5)
        ws2.Cells(row, 16) = tmpA(6) + tmpA(7) + tmpA(8)
        ws2.Cells(row, 17) = tmpA(9) + tmpA(10) + tmpA(11)
        ws2.Cells(row, 18) = mTotal
        ws2.Cells(row, 19) = Now()
        ws2.Cells(row, 19).NumberFormat = "dd/mm/yyyy hh:mm"

        If mTotal > 30 Then
            ws2.Cells(row, 18).Font.Color = RGB(204, 0, 0)
            ws2.Cells(row, 18).Font.Bold = True
        End If

        row = row + 1
    Next i

    FormatTable ws2, row - 1, 19
    ws2.Columns("A").ColumnWidth = 28
    ws2.Range("B1:S1").ColumnWidth = 8

    ' === Sheet 3: Equite WE ===
    Dim ws3 As Worksheet
    Set ws3 = GetOrCreateSheet("Equite WE")

    Dim h3 As Variant
    h3 = Array("Agent", "Samedis", "Dimanches", "Total WE", "Ecart Moyenne", "Derniere MAJ")
    For c = 0 To UBound(h3)
        ws3.Cells(1, c + 1).value = h3(c)
    Next c
    FormatHeader ws3, 6

    Dim totalWE As Double: totalWE = 0
    For i = 0 To agentCount - 1
        Dim sArr() As Double, dArr() As Double
        sArr = dictWE_S(agentOrder(i))
        dArr = dictWE_D(agentOrder(i))
        totalWE = totalWE + SumArr(sArr) + SumArr(dArr)
    Next i
    Dim avgWE As Double
    If agentCount > 0 Then avgWE = totalWE / agentCount Else avgWE = 0

    row = 2
    For i = 0 To agentCount - 1
        agName = agentOrder(i)
        ws3.Cells(row, 1).value = agName

        sArr = dictWE_S(agName): dArr = dictWE_D(agName)
        Dim sTotal As Double, dTotal As Double, weTotal As Double
        sTotal = SumArr(sArr): dTotal = SumArr(dArr): weTotal = sTotal + dTotal

        ws3.Cells(row, 2) = sTotal
        ws3.Cells(row, 3) = dTotal
        ws3.Cells(row, 4) = weTotal
        ws3.Cells(row, 4).Font.Bold = True
        ws3.Cells(row, 5) = Round(weTotal - avgWE, 1)
        ws3.Cells(row, 6) = Now()
        ws3.Cells(row, 6).NumberFormat = "dd/mm/yyyy hh:mm"

        If Abs(weTotal - avgWE) > 10 Then
            ws3.Cells(row, 5).Font.Color = RGB(204, 0, 0)
            ws3.Cells(row, 5).Font.Bold = True
        End If

        row = row + 1
    Next i

    FormatTable ws3, row - 1, 6
    ws3.Columns("A").ColumnWidth = 28
    ws3.Range("B1:F1").ColumnWidth = 14

    ' === Sheet 4: Detail Agent ===
    Dim ws4 As Worksheet
    Set ws4 = GetOrCreateSheet("Detail Agent")

    Dim h4 As Variant
    h4 = Array("Agent", "Mois", "CA", "EL", "ANC", "C SOC", "DP", "CRP", "Maladie", "Jours Trav.", "WE Trav.", "Derniere MAJ")
    For c = 0 To UBound(h4)
        ws4.Cells(1, c + 1).value = h4(c)
    Next c
    FormatHeader ws4, UBound(h4) + 1

    Dim mNamesArr() As String
    mNamesArr = Split("Jan,Fev,Mar,Avr,Mai,Jun,Jul,Aou,Sep,Oct,Nov,Dec", ",")

    row = 2
    For i = 0 To agentCount - 1
        agName = agentOrder(i)
        For m = 0 To 11
            ws4.Cells(row, 1).value = agName
            ws4.Cells(row, 2).value = mNamesArr(m)

            Dim tCA() As Double, tEL() As Double, tANC() As Double
            Dim tCS() As Double, tDP2() As Double, tCR() As Double
            Dim tML() As Double, tWK() As Double

            tCA = dictCA(agName): ws4.Cells(row, 3) = tCA(m)
            tEL = dictEL(agName): ws4.Cells(row, 4) = tEL(m)
            tANC = dictANC(agName): ws4.Cells(row, 5) = tANC(m)
            tCS = dictCSOC(agName): ws4.Cells(row, 6) = tCS(m)
            tDP2 = dictDP(agName): ws4.Cells(row, 7) = tDP2(m)
            tCR = dictCRP(agName): ws4.Cells(row, 8) = tCR(m)
            tML = dictMAL(agName): ws4.Cells(row, 9) = tML(m)
            tWK = dictWORK(agName): ws4.Cells(row, 10) = tWK(m)

            sArr = dictWE_S(agName): dArr = dictWE_D(agName)
            ws4.Cells(row, 11) = sArr(m) + dArr(m)
            ws4.Cells(row, 12) = Now()
            ws4.Cells(row, 12).NumberFormat = "dd/mm/yyyy hh:mm"

            row = row + 1
        Next m
    Next i

    FormatTable ws4, row - 1, 12
    ws4.Columns("A").ColumnWidth = 28
    ws4.Columns("B").ColumnWidth = 8
    ws4.Range("C1:L1").ColumnWidth = 10
    ws4.Range("C2").Select
    ActiveWindow.FreezePanes = True

    ' Activer le premier onglet
    ws1.Activate

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Suivi RH genere avec succes !" & vbCrLf & vbCrLf & _
        "4 onglets mis a jour :" & vbCrLf & _
        "  - Soldes Conges (CA/EL/ANC/C SOC/DP/CRP1)" & vbCrLf & _
        "  - Absenteisme (maladie par mois)" & vbCrLf & _
        "  - Equite WE (samedis/dimanches)" & vbCrLf & _
        "  - Detail Agent (mensuel)" & vbCrLf & vbCrLf & _
        "(v2: quotas depuis Config_Personnel, onglets permanents)", vbInformation, "CTR Suivi RH 2026"

    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur: " & Err.Description & " (ligne " & Erl & ")", vbCritical, "Erreur Suivi RH"
End Sub

' ---- HELPERS ----
Private Function SumArr(arr() As Double) As Double
    Dim s As Double, j As Long
    For j = LBound(arr) To UBound(arr)
        s = s + arr(j)
    Next j
    SumArr = s
End Function
