Attribute VB_Name = "Module_JoursFeries"

Option Explicit

'===================================================================================
' Regenerer les codes F/R des jours feries a partir de l'annee de Feuil_Config
'===================================================================================
Public Sub MettreAJourConfigurationCodes()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet, wsCfg As Worksheet
    Set ws = ThisWorkbook.Sheets("Config_Codes")
    Set wsCfg = ThisWorkbook.Sheets("Feuil_Config")

    Dim annee As Long
    annee = GetCfgYear(wsCfg)
    If annee < 1900 Or annee > 2100 Then
        MsgBox "Annee invalide dans Feuil_Config (CFG_Year ou AnneePlanning).", vbCritical
        GoTo CleanExit
    End If

    Dim colCode As Long, colDesc As Long, colType As Long, colHN As Long
    Dim colTop As Long, colHS As Long, colHPS As Long, colHPE As Long, colHE As Long
    Dim colF645 As Long, colF78 As Long, colMatin As Long, colPM As Long, colSoir As Long, colNuit As Long

    colCode = FindHeaderCol(ws, "Code")
    colDesc = FindHeaderCol(ws, "Description")
    colType = FindHeaderCol(ws, "Type_Code")
    colHN = FindHeaderCol(ws, "Heures_normales")
    colTop = FindHeaderCol(ws, "TopCode")
    colHS = FindHeaderCol(ws, "H_Start")
    colHPS = FindHeaderCol(ws, "H_Pause_Start")
    colHPE = FindHeaderCol(ws, "H_Pause_End")
    colHE = FindHeaderCol(ws, "H_End")
    colF645 = FindHeaderCol(ws, "F_6h45")
    colF78 = FindHeaderCol(ws, "F_7h_8h")
    colMatin = FindHeaderCol(ws, "Matin")
    colPM = FindHeaderCol(ws, "PM")
    colSoir = FindHeaderCol(ws, "Soir")
    colNuit = FindHeaderCol(ws, "Nuit")

    If colCode = 0 Then
        MsgBox "Colonne 'Code' introuvable dans Config_Codes.", vbCritical
        GoTo CleanExit
    End If

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.count, colCode).End(xlUp).row

    ' Supprimer anciens codes F/R
    For r = lastRow To 2 Step -1
        Dim code As String
        code = Trim(CStr(ws.Cells(r, colCode).value))
        If (Left$(code, 2) = "F " Or Left$(code, 2) = "R ") And InStr(code, "-") > 0 Then
            ws.Rows(r).Delete
        End If
    Next r

    ' Construire liste de dates feries
    Dim feries As Variant
    feries = BuildFeriesBE(annee)

    Dim nb As Long
    nb = UBound(feries) - LBound(feries) + 1
    If nb <= 0 Then GoTo CleanExit

    ' Inserer lignes
    ws.Rows(2).Resize(nb * 2).Insert

    ' Remplir lignes
    Dim i As Long, rowIdx As Long, d As Date
    rowIdx = 2
    For i = LBound(feries) To UBound(feries)
        d = feries(i)
        Call WriteFerieRow(ws, rowIdx, colCode, colDesc, colType, colHN, colTop, colHS, colHPS, colHPE, colHE, colF645, colF78, colMatin, colPM, colSoir, colNuit, "F ", d)
        rowIdx = rowIdx + 1
        Call WriteFerieRow(ws, rowIdx, colCode, colDesc, colType, colHN, colTop, colHS, colHPS, colHPE, colHE, colF645, colF78, colMatin, colPM, colSoir, colNuit, "R ", d)
        rowIdx = rowIdx + 1
    Next i

    MsgBox "Config_Codes mis a jour (F/R) pour l'annee " & annee & ".", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "Erreur MettreAJourConfigurationCodes: " & Err.Number & " - " & Err.description, vbCritical
    Resume CleanExit
End Sub

Private Sub WriteFerieRow(ws As Worksheet, rowIdx As Long, _
    colCode As Long, colDesc As Long, colType As Long, colHN As Long, colTop As Long, _
    colHS As Long, colHPS As Long, colHPE As Long, colHE As Long, colF645 As Long, colF78 As Long, _
    colMatin As Long, colPM As Long, colSoir As Long, colNuit As Long, _
    prefix As String, d As Date)

    Dim code As String
    code = prefix & Day(d) & "-" & Month(d)

    ws.Cells(rowIdx, colCode).value = code
    If colDesc > 0 Then ws.Cells(rowIdx, colDesc).value = "Férié"
    If colType > 0 Then ws.Cells(rowIdx, colType).value = "Férié"
    If colHN > 0 Then ws.Cells(rowIdx, colHN).value = 0

    If colTop > 0 Then ws.Cells(rowIdx, colTop).value = ""
    If colHS > 0 Then ws.Cells(rowIdx, colHS).value = ""
    If colHPS > 0 Then ws.Cells(rowIdx, colHPS).value = ""
    If colHPE > 0 Then ws.Cells(rowIdx, colHPE).value = ""
    If colHE > 0 Then ws.Cells(rowIdx, colHE).value = ""

    If colF645 > 0 Then ws.Cells(rowIdx, colF645).value = 0
    If colF78 > 0 Then ws.Cells(rowIdx, colF78).value = 0
    If colMatin > 0 Then ws.Cells(rowIdx, colMatin).value = 0
    If colPM > 0 Then ws.Cells(rowIdx, colPM).value = 0
    If colSoir > 0 Then ws.Cells(rowIdx, colSoir).value = 0
    If colNuit > 0 Then ws.Cells(rowIdx, colNuit).value = 0
End Sub

Private Function FindHeaderCol(ws As Worksheet, headerName As String) As Long
    Dim c As Long, v As String
    For c = 1 To ws.UsedRange.Columns.count
        v = Trim(CStr(ws.Cells(1, c).value))
        If StrComp(v, headerName, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Private Function GetCfgYear(wsCfg As Worksheet) As Long
    Dim lastRow As Long, r As Long
    Dim key As String
    lastRow = wsCfg.Cells(wsCfg.Rows.count, "A").End(xlUp).row

    For r = 2 To lastRow
        key = Trim(CStr(wsCfg.Cells(r, 1).value))
        If StrComp(key, "CFG_Year", vbTextCompare) = 0 Then
            GetCfgYear = CLng(wsCfg.Cells(r, 2).value)
            Exit Function
        End If
    Next r

    For r = 2 To lastRow
        key = Trim(CStr(wsCfg.Cells(r, 1).value))
        If StrComp(key, "AnneePlanning", vbTextCompare) = 0 Then
            GetCfgYear = CLng(wsCfg.Cells(r, 2).value)
            Exit Function
        End If
    Next r

    GetCfgYear = Year(Date)
End Function

'===================================================================================
' JOURS FERIES (BE)
'===================================================================================
Private Function BuildFeriesBE(ByVal annee As Long) As Variant
    Dim paques As Date
    paques = CalculerPaques(annee)

    Dim arr(1 To 10) As Date
    arr(1) = DateSerial(annee, 1, 1)
    arr(2) = paques + 1
    arr(3) = DateSerial(annee, 5, 1)
    arr(4) = paques + 39
    arr(5) = paques + 50
    arr(6) = DateSerial(annee, 7, 21)
    arr(7) = DateSerial(annee, 8, 15)
    arr(8) = DateSerial(annee, 11, 1)
    arr(9) = DateSerial(annee, 11, 11)
    arr(10) = DateSerial(annee, 12, 25)

    ' Tri simple (10 elements)
    Dim i As Long, j As Long, tmp As Date
    For i = 1 To 9
        For j = i + 1 To 10
            If arr(j) < arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    BuildFeriesBE = arr
End Function

Private Function CalculerPaques(ByVal annee As Long) As Date
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


