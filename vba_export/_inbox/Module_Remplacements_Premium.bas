' ExportedAt: 2026-01-13 14:15:00 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module_Remplacements_Premium"
Option Explicit

'================================================================================================
'  CONFIG ACCESS (LOCAL, NO DEPENDENCIES)
'  Source: Sheet "Feuil_Config" (A=Key, B=Value)
'================================================================================================

Private Const CFG_SHEET As String = "Feuil_Config"
Private Const CFG_KEY_COL As Long = 1   'A
Private Const CFG_VAL_COL As Long = 2   'B

' Cache (key -> value) to speed up repeated reads
Private mCfgCache As Object
Private mCfgCacheBuilt As Boolean

Private Sub CfgCache_Ensure()
    If mCfgCacheBuilt Then Exit Sub
    Set mCfgCache = CreateObject("Scripting.Dictionary")
    mCfgCache.CompareMode = 1 ' vbTextCompare

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim k As String, v As String

    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, CFG_KEY_COL).End(xlUp).Row

    For r = 1 To lastRow
        k = Trim$(CStr(ws.Cells(r, CFG_KEY_COL).Value2))
        If Len(k) > 0 Then
            v = Trim$(CStr(ws.Cells(r, CFG_VAL_COL).Value2))
            ' keep first occurrence (avoid surprises if duplicates)
            If Not mCfgCache.Exists(k) Then mCfgCache.Add k, v
        End If
    Next r

    mCfgCacheBuilt = True
    Exit Sub

EH:
    ' If Feuil_Config missing -> keep empty cache (no crash)
    Set mCfgCache = CreateObject("Scripting.Dictionary")
    mCfgCache.CompareMode = 1
    mCfgCacheBuilt = True
End Sub

' Call this if you change Feuil_Config programmatically and want fresh values.
Public Sub CfgCache_Reset()
    mCfgCacheBuilt = False
    Set mCfgCache = Nothing
End Sub

' Returns raw string ("" if missing)
Private Function CfgGetRaw(ByVal key As String) As String
    CfgCache_Ensure
    key = Trim$(key)
    If Len(key) = 0 Then
        CfgGetRaw = vbNullString
    ElseIf mCfgCache.Exists(key) Then
        CfgGetRaw = CStr(mCfgCache(key))
    Else
        CfgGetRaw = vbNullString
    End If
End Function

' Text with default
' Variant with default + best-effort conversion based on defaultVal type
' ===================================================================================
' MODULE :          Module_Remplacements_Premium
' DESCRIPTION :     Version optimisee et config-driven de la generation de demandes
'                   de remplacement. Lit tous les parametres depuis tblCFG.
' ===================================================================================

' --- Constantes pour lisibilite ---
Private Const COL_FIRST_DAY As Long = 3   ' Colonne C
Private Const COL_LAST_DAY As Long = 33   ' Colonne AG
Private Const ROW_START_DATES As Long = 7 ' Premiere ligne des dates dans le modele

' ===================================================================================
' MACRO PRINCIPALE
' ===================================================================================

' Wrapper pour compatibilite avec l'ancien nom
Public Sub GenerateNewWorkbookAndFillDates_Optimized_V4(nomPrenom As String, _
                                                        dayOrNight As String, _
                                                        postCM As String, _
                                                        ReplacementLines As String)
    GenerateNewWorkbookAndFillDates_Premium nomPrenom, dayOrNight, postCM, ReplacementLines
End Sub

Public Sub GenerateNewWorkbookAndFillDates_Premium(nomPrenom As String, _
                                                    dayOrNight As String, _
                                                    postCM As String, _
                                                    ReplacementLines As String)
    On Error GoTo ErrorHandler
    
    ' --- Variables principales ---
    Dim wb As Workbook, newWb As Workbook
    Dim wsSource As Worksheet, wsModel As Worksheet, wsFinal As Worksheet
    Dim yearToUse As Long, monthNumber As Long
    Dim linesArray() As String, mappedLines() As Long
    Dim demandsByLine As Object
    Dim sheetCount As Long
    
    ' --- Parametres depuis config ---
    Dim lineOffset As Long
    Dim asbdColor As Long
    Dim nurseCodes As String
    Dim holidayPrefixes As String
    Dim savePath As String
    Dim holidaySheetName As String
    Dim dictHolidays As Object
    
    ' --- Optimisation performance ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Set wb = ThisWorkbook
    Set wsSource = wb.ActiveSheet
    
    ' =================================================================
    ' ETAPE 1: CHARGER CONFIGURATION
    ' =================================================================
    If Not LoadConfig(lineOffset, asbdColor, nurseCodes, holidayPrefixes, savePath, holidaySheetName) Then
        GoTo CleanUp
    End If
    
    ' --- Determiner annee et mois ---
    yearToUse = GetDynamicYear()
    monthNumber = GetMonthFromSheetName(wsSource.Name)
    If monthNumber = 0 Then
        MsgBox "L'onglet '" & wsSource.Name & "' n'est pas un mois valide.", vbExclamation
        GoTo CleanUp
    End If
    
    ' --- Charger jours feries ---
    Set dictHolidays = LoadHolidays(yearToUse, holidayPrefixes, holidaySheetName)
    
    ' --- Feuille modele ---
    On Error Resume Next
    Set wsModel = wb.Sheets("Model")
    On Error GoTo ErrorHandler
    If wsModel Is Nothing Then
        MsgBox "Feuille 'Model' introuvable!", vbCritical
        GoTo CleanUp
    End If
    
    ' =================================================================
    ' ETAPE 2: VALIDER ET MAPPER LIGNES
    ' =================================================================
    If Trim(ReplacementLines) = "" Then
        MsgBox "Aucune ligne de remplacement specifiee.", vbExclamation
        GoTo CleanUp
    End If
    
    If Not ParseAndMapLines(ReplacementLines, lineOffset, linesArray, mappedLines) Then
        GoTo CleanUp
    End If
    
    ' =================================================================
    ' ETAPE 3: COLLECTER DEMANDES
    ' =================================================================
    Set demandsByLine = CollectDemands(wsSource, mappedLines, yearToUse, monthNumber, _
                                       asbdColor, nurseCodes, dictHolidays)
    
    If demandsByLine.Count = 0 Then
        MsgBox "Aucune demande valide trouvee pour les lignes specifiees.", vbInformation
        GoTo CleanUp
    End If
    
    ' =================================================================
    ' ETAPE 4: CREER CLASSEUR ET FEUILLES
    ' =================================================================
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    
    sheetCount = CreateDemandSheets(newWb, wsModel, demandsByLine, linesArray, mappedLines, _
                                    lineOffset, yearToUse, monthNumber, asbdColor)
    
    If sheetCount = 0 Then
        MsgBox "Aucune feuille n'a pu etre generee.", vbExclamation
        newWb.Close SaveChanges:=False
        GoTo CleanUp
    End If
    
    ' Supprimer feuille par defaut
    On Error Resume Next
    If newWb.Sheets.Count > sheetCount Then newWb.Sheets(1).Delete
    On Error GoTo ErrorHandler
    
    ' =================================================================
    ' ETAPE 5: FORMATAGE FINAL ET SAUVEGARDE
    ' =================================================================
    FormatAndColorize newWb, dictHolidays, asbdColor, yearToUse, monthNumber
    
    SaveWorkbook newWb, postCM, nomPrenom, dayOrNight, yearToUse, monthNumber, savePath, sheetCount
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    MsgBox "Erreur #" & Err.Number & ": " & Err.Description, vbCritical, "Erreur"
    Resume CleanUp
End Sub

' ===================================================================================
' FONCTIONS DE CHARGEMENT CONFIGURATION
' ===================================================================================

Private Function LoadConfig(ByRef lineOffset As Long, ByRef asbdColor As Long, _
                           ByRef nurseCodes As String, ByRef holidayPrefixes As String, _
                           ByRef savePath As String, ByRef holidaySheetName As String) As Boolean
    LoadConfig = False
    
    On Error GoTo ConfigError
    
    lineOffset = CLng(CfgValueOr("DecalageLigneRemplacement", 0))
    asbdColor = CLng(CfgValueOr("Couleur_ASBD_RGB", 16777215))
    nurseCodes = CStr(CfgTextOr("CodesInfirmiere", "*;INF;IDE;IC"))
    holidayPrefixes = CStr(CfgTextOr("Prefixe_JourFerie", "JF;FERIE"))
    holidaySheetName = CStr(CfgTextOr("OngletJoursFeries", "Config_Calendrier"))
    
    savePath = CStr(CfgTextOr("CheminSauvegarde", ""))
    savePath = Replace(savePath, "{annee}", GetDynamicYear())
    savePath = Replace(savePath, "{username}", Environ("USERNAME"))
    savePath = Replace(savePath, "{workbook}", GetLocalWorkbookPath())
    
    ' Note: La creation du dossier est geree par SaveWorkbook avec dialogue
    
    LoadConfig = True
    Exit Function

ConfigError:
    MsgBox "Erreur chargement config: " & Err.Description, vbCritical
End Function

' ===================================================================================
' FONCTIONS DE PARSING
' ===================================================================================

Private Function ParseAndMapLines(ByVal inputLines As String, ByVal offset As Long, _
                                  ByRef linesArray() As String, ByRef mappedLines() As Long) As Boolean
    ParseAndMapLines = False
    Dim i As Long
    
    linesArray = Split(inputLines, ",")
    ReDim mappedLines(LBound(linesArray) To UBound(linesArray))
    
    For i = LBound(linesArray) To UBound(linesArray)
        If Not IsNumeric(Trim(linesArray(i))) Then
            MsgBox "Ligne non numerique: '" & linesArray(i) & "'.", vbCritical
            Exit Function
        End If
        mappedLines(i) = CLng(Trim(linesArray(i))) + offset
    Next i
    
    ParseAndMapLines = True
End Function

Private Function GetMonthFromSheetName(sheetName As String) As Long
    Dim months As Object
    Set months = CreateObject("Scripting.Dictionary")
    months.CompareMode = vbTextCompare
    
    months.Add "janv", 1: months.Add "janvier", 1
    months.Add "fev", 2: months.Add "fevrier", 2: months.Add "f" & ChrW(233) & "vrier", 2
    months.Add "mars", 3
    months.Add "avril", 4
    months.Add "mai", 5
    months.Add "juin", 6
    months.Add "juil", 7: months.Add "juillet", 7
    months.Add "aout", 8: months.Add "ao" & ChrW(251) & "t", 8
    months.Add "sept", 9: months.Add "septembre", 9
    months.Add "oct", 10: months.Add "octobre", 10
    months.Add "nov", 11: months.Add "novembre", 11
    months.Add "dec", 12: months.Add "decembre", 12: months.Add "d" & ChrW(233) & "cembre", 12
    
    If months.Exists(LCase(sheetName)) Then
        GetMonthFromSheetName = months(LCase(sheetName))
    Else
        GetMonthFromSheetName = 0
    End If
End Function

' ===================================================================================
' FONCTIONS DE COLLECTE DES DEMANDES
' ===================================================================================

Private Function CollectDemands(wsSource As Worksheet, mappedLines() As Long, _
                                yearToUse As Long, monthNumber As Long, _
                                asbdColor As Long, nurseCodes As String, _
                                dictHolidays As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, j As Long, lineNum As Long
    Dim sourceValues As Variant, cellValue As Variant
    Dim sourceCell As Range
    Dim demand As CReplacementInfo
    Dim tempDate As Date
    Dim lineDemands As Collection
    
    For i = LBound(mappedLines) To UBound(mappedLines)
        lineNum = mappedLines(i)
        
        sourceValues = wsSource.Range(wsSource.Cells(lineNum, COL_FIRST_DAY), _
                                      wsSource.Cells(lineNum, COL_LAST_DAY)).Value
        
        If IsArray(sourceValues) Then
            For j = 1 To UBound(sourceValues, 2)
                cellValue = sourceValues(1, j)
                
                If Not IsEmpty(cellValue) And Trim(CStr(cellValue)) <> "" Then
                    Set demand = New CReplacementInfo
                    demand.sourceLineNum = lineNum
                    demand.shiftCode = Trim(CStr(cellValue))
                    
                    Set sourceCell = wsSource.Cells(lineNum, j + COL_FIRST_DAY - 1)
                    demand.IsASBD = (sourceCell.Interior.Color = asbdColor)
                    demand.IsNurse = IsNurseShift(demand.shiftCode, nurseCodes)
                    
                    On Error Resume Next
                    tempDate = DateSerial(yearToUse, monthNumber, j)
                    If Err.Number = 0 Then
                        demand.DemandDate = tempDate
                        demand.IsWeekend = (Weekday(demand.DemandDate, vbMonday) >= 6)
                        demand.IsHoliday = dictHolidays.Exists(CLng(demand.DemandDate))
                        
                        If Not result.Exists(lineNum) Then
                            Set result(lineNum) = New Collection
                        End If
                        result(lineNum).Add demand
                    End If
                    On Error GoTo 0
                End If
            Next j
        End If
    Next i
    
    Set CollectDemands = result
End Function

Private Function IsNurseShift(shiftCode As String, configString As String) As Boolean
    IsNurseShift = False
    If Trim(configString) = "" Or Trim(shiftCode) = "" Then Exit Function
    
    Dim rules() As String, rule As Variant
    rules = Split(configString, ";")
    
    For Each rule In rules
        If Trim(CStr(rule)) = "*" Then
            If InStr(1, shiftCode, "*", vbTextCompare) > 0 Then
                IsNurseShift = True
                Exit Function
            End If
        Else
            If UCase(Trim(shiftCode)) = UCase(Trim(CStr(rule))) Then
                IsNurseShift = True
                Exit Function
            End If
        End If
    Next rule
End Function

' ===================================================================================
' FONCTIONS DE CREATION DES FEUILLES
' ===================================================================================

Private Function CreateDemandSheets(newWb As Workbook, wsModel As Worksheet, _
                                    demandsByLine As Object, linesArray() As String, _
                                    mappedLines() As Long, lineOffset As Long, _
                                    yearToUse As Long, monthNumber As Long, _
                                    asbdColor As Long) As Long
    Dim sheetCount As Long
    Dim lineKey As Variant
    Dim lineDemands As Collection
    Dim originalNum As Long
    Dim hasAS As Boolean, hasINF As Boolean
    Dim d As CReplacementInfo
    Dim wsFinal As Worksheet
    Dim transposedRow As Long
    
    sheetCount = 0
    
    For Each lineKey In demandsByLine.Keys
        Set lineDemands = demandsByLine(lineKey)
        originalNum = GetOriginalLineNum(CLng(lineKey), mappedLines, linesArray, lineOffset)
        
        ' Determiner types de demandes
        hasAS = False: hasINF = False
        For Each d In lineDemands
            If d.IsNurse Then hasINF = True Else hasAS = True
            If hasAS And hasINF Then Exit For
        Next d
        
        ' Creer feuille AS
        If hasAS Then
            Set wsFinal = CreateSheet(newWb, wsModel, "AS", CLng(lineKey), originalNum, monthNumber, yearToUse)
            If Not wsFinal Is Nothing Then
                sheetCount = sheetCount + 1
                For Each d In lineDemands
                    If Not d.IsNurse Then
                        transposedRow = Day(d.DemandDate) + ROW_START_DATES - 1
                        WriteDemand wsFinal, transposedRow, d, asbdColor
                    End If
                Next d
            End If
        End If
        
        ' Creer feuille INF
        If hasINF Then
            Set wsFinal = CreateSheet(newWb, wsModel, "INF", CLng(lineKey), originalNum, monthNumber, yearToUse)
            If Not wsFinal Is Nothing Then
                sheetCount = sheetCount + 1
                For Each d In lineDemands
                    If d.IsNurse Then
                        transposedRow = Day(d.DemandDate) + ROW_START_DATES - 1
                        WriteDemand wsFinal, transposedRow, d, asbdColor
                    End If
                Next d
            End If
        End If
    Next lineKey
    
    CreateDemandSheets = sheetCount
End Function

Private Function CreateSheet(newWb As Workbook, wsModel As Worksheet, _
                             sheetType As String, sourceLineNum As Long, _
                             originalNum As Long, monthNum As Long, yearNum As Long) As Worksheet
    Dim wsNew As Worksheet
    Dim sheetName As String, titleText As String
    
    On Error GoTo SheetError
    
    wsModel.Copy After:=newWb.Sheets(newWb.Sheets.Count)
    Set wsNew = newWb.Sheets(newWb.Sheets.Count)
    
    ' Nom et titre
    If sheetType = "AS" Then
        titleText = "Aide-Soignant (Source L" & sourceLineNum & ")"
        If originalNum <> -1 Then
            sheetName = "AS_Rempl_" & originalNum
        Else
            sheetName = "AS_Ligne_" & sourceLineNum
        End If
    Else
        titleText = "Infirmiere (Source L" & sourceLineNum & ")"
        If originalNum = 6 Then
            sheetName = "INF_Nuit_1"
        ElseIf originalNum = 7 Then
            sheetName = "INF_Nuit_2"
        ElseIf originalNum <> -1 Then
            sheetName = "INF_Rempl_" & originalNum
        Else
            sheetName = "INF_Ligne_" & sourceLineNum
        End If
    End If
    
    wsNew.Name = CleanSheetName(sheetName, newWb)
    wsNew.Activate
    ActiveWindow.Zoom = 70
    wsNew.Range("F3").Value = titleText
    
    ' Remplir dates
    FillDates wsNew, monthNum, yearNum
    
    Set CreateSheet = wsNew
    Exit Function

SheetError:
    Set CreateSheet = Nothing
End Function

Private Sub WriteDemand(ws As Worksheet, targetRow As Long, d As CReplacementInfo, asbdColor As Long)
    With ws.Cells(targetRow, 3)
        .Value = ConvertShiftCode(d.shiftCode)
        .HorizontalAlignment = xlCenter
        .Font.Name = "Arial Narrow"
        .Font.Size = 16
        
        If d.IsNurse Then .Value = Replace(.Value, "*", "")
        If d.IsASBD Then .Interior.Color = asbdColor Else .Interior.ColorIndex = xlNone
    End With
    
    If d.IsNurse Then ws.Cells(targetRow, 5).Value = "Infirmiere"
End Sub

Private Sub FillDates(ws As Worksheet, monthNum As Long, yearNum As Long)
    Dim daysInMonth As Long, startDate As Date, i As Long
    Dim datesArr(), daysArr()
    
    daysInMonth = Day(DateSerial(yearNum, monthNum + 1, 0))
    startDate = DateSerial(yearNum, monthNum, 1)
    
    ReDim datesArr(1 To daysInMonth, 1 To 1)
    ReDim daysArr(1 To daysInMonth, 1 To 1)
    
    For i = 1 To daysInMonth
        datesArr(i, 1) = startDate + i - 1
        daysArr(i, 1) = StrConv(Format(startDate + i - 1, "ddd"), vbProperCase)
    Next i
    
    With ws.Range("A" & ROW_START_DATES).Resize(daysInMonth, 2)
        .ClearContents
        .Columns(1).Value = daysArr
        .Columns(2).Value = datesArr
        .Columns(2).NumberFormat = "dd/mm/yyyy"
    End With
End Sub

' ===================================================================================
' FONCTIONS DE FORMATAGE
' ===================================================================================

Private Sub FormatAndColorize(newWb As Workbook, holidays As Object, asbdColor As Long, _
                              yearNum As Long, monthNum As Long)
    Dim ws As Worksheet, iRow As Long
    Dim lastRow As Long, dateVal As Variant
    Dim rng As Range
    Dim colorWeekend As Long, colorFerie As Long
    Dim colorPoliceWE As Long, colorPoliceFerie As Long
    
    ' Lire couleurs depuis config
    colorWeekend = CLng(CfgValueOr("REMPL_Couleur_Weekend", 65535))        ' Jaune
    colorFerie = CLng(CfgValueOr("REMPL_Couleur_Ferie", 255))              ' Rouge
    colorPoliceWE = CLng(CfgValueOr("REMPL_Couleur_Police_Weekend", 16711680))  ' Bleu
    colorPoliceFerie = CLng(CfgValueOr("REMPL_Couleur_Police_Ferie", 0))   ' Noir
    
    lastRow = ROW_START_DATES + Day(DateSerial(yearNum, monthNum + 1, 0)) - 1
    
    For Each ws In newWb.Worksheets
        ws.Range("A6:F" & lastRow).Borders.LineStyle = xlContinuous
        
        For iRow = ROW_START_DATES To lastRow
            If IsDate(ws.Cells(iRow, 2).Value) Then
                dateVal = ws.Cells(iRow, 2).Value
                Set rng = ws.Range("A" & iRow & ":C" & iRow)
                
                If holidays.Exists(CLng(CDate(dateVal))) Then
                    ' Jour ferie
                    rng.Interior.Color = colorFerie
                    rng.Font.Bold = True
                    rng.Font.Color = colorPoliceFerie
                ElseIf Weekday(CDate(dateVal), vbMonday) >= 6 Then
                    ' Weekend
                    rng.Interior.Color = colorWeekend
                    rng.Font.Bold = True
                    rng.Font.Color = colorPoliceWE
                Else
                    ' Jour normal
                    If ws.Cells(iRow, 3).Interior.Color <> asbdColor Then
                        rng.Interior.ColorIndex = xlNone
                    End If
                    rng.Font.Bold = False
                    rng.Font.Color = 0
                End If
            End If
        Next iRow
    Next ws
End Sub

' ===================================================================================
' FONCTIONS DE SAUVEGARDE
' ===================================================================================

Private Sub SaveWorkbook(wbToSave As Workbook, postCM As String, nomPrenom As String, _
                        dayOrNight As String, yearNum As Long, monthNum As Long, _
                        savePath As String, sheetCount As Long)
    Dim baseFileName As String, finalName As String
    Dim cloudUrl As String
    Dim fullPath As String
    
    ' Construire le nom du fichier
    If UCase(Trim(postCM)) = "/ MOIS" Then
        baseFileName = "Demandes_" & Format(DateSerial(yearNum, monthNum, 1), "mmmm_yyyy")
    ElseIf UCase(Trim(postCM)) Like "US*1D*" Then
        baseFileName = "Demande_Us_1D_" & yearNum & "-" & Format(monthNum, "00")
    Else
        baseFileName = "Demande_" & Format(Date, "yyyy-mm-dd") & "_" & Format(Time, "hhmmss")
    End If
    
    finalName = CleanFileName(baseFileName & ".xlsm")
    
    ' Si chemin vide ou invalide, utiliser Documents
    If Trim(savePath) = "" Or Left(savePath, 4) = "http" Then
        savePath = Environ("USERPROFILE") & "\Documents\DemandeRemplacements"
    End If
    
    ' Creer dossier si necessaire
    On Error Resume Next
    If Dir(savePath, vbDirectory) = "" Then
        MkDir savePath
    End If
    On Error GoTo 0
    
    fullPath = savePath & "\" & finalName
    
    ' Tentative de sauvegarde
    On Error Resume Next
    wbToSave.SaveAs FileName:=fullPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    If Err.Number <> 0 Then
        ' Echec - proposer SaveAs dialog
        Err.Clear
        On Error GoTo 0
        
        MsgBox "Impossible de sauvegarder automatiquement." & vbCrLf & _
               "Chemin: " & fullPath & vbCrLf & vbCrLf & _
               "Un dialogue va s'ouvrir pour choisir l'emplacement.", vbExclamation
        
        Application.DisplayAlerts = True
        wbToSave.Activate
        
        Dim result As Boolean
        result = Application.Dialogs(xlDialogSaveAs).Show(finalName)
        
        If Not result Then
            MsgBox "Sauvegarde annulee.", vbExclamation
            Exit Sub
        End If
    End If
    On Error GoTo 0
    
    ' Message de succes
    cloudUrl = CfgTextOr("CheminSauvegarde_Cloud", "")
    
    If cloudUrl <> "" Then
        MsgBox "Classeur cree avec " & sheetCount & " feuille(s)." & vbCrLf & vbCrLf & _
               "Emplacement: " & wbToSave.FullName & vbCrLf & vbCrLf & _
               "Lien OneDrive:" & vbCrLf & cloudUrl, vbInformation, "Succes"
    Else
        MsgBox "Classeur cree avec " & sheetCount & " feuille(s)." & vbCrLf & _
               "Emplacement: " & wbToSave.FullName, vbInformation
    End If
End Sub

Private Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

' ===================================================================================
' FONCTIONS UTILITAIRES
' ===================================================================================

Private Function LoadHolidays(yearToUse As Long, prefixes As String, sheetName As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim data As Variant, cellValue As String
    Dim prefixArr() As String, prefix As Variant
    Dim datePart As String, parts() As String
    Dim dayNum As Long, monthNum As Long, dt As Date
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set LoadHolidays = dict
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        Set LoadHolidays = dict
        Exit Function
    End If
    
    data = ws.Range("A2:A" & lastRow).Value
    prefixArr = Split(prefixes, ";")
    
    If IsArray(data) Then
        For r = LBound(data, 1) To UBound(data, 1)
            cellValue = Trim(CStr(data(r, 1)))
            For Each prefix In prefixArr
                If UCase(Left(cellValue, Len(Trim(CStr(prefix))))) = UCase(Trim(CStr(prefix))) Then
                    datePart = Trim(Mid(cellValue, Len(Trim(CStr(prefix))) + 1))
                    parts = Split(datePart, "-")
                    If UBound(parts) >= 1 Then
                        On Error Resume Next
                        dayNum = CLng(Trim(parts(0)))
                        monthNum = CLng(Trim(parts(1)))
                        dt = DateSerial(yearToUse, monthNum, dayNum)
                        If Err.Number = 0 And Year(dt) = yearToUse Then
                            dict(CLng(dt)) = True
                        End If
                        Err.Clear
                        On Error GoTo 0
                    End If
                    Exit For
                End If
            Next prefix
        Next r
    End If
    
    Set LoadHolidays = dict
End Function

Private Function GetDynamicYear() As Long
    Dim y As Variant
    y = CfgValueOr("ANNEE_PLANNING", 0)
    If IsNumeric(y) Then
        If CLng(y) >= 1900 And CLng(y) <= 2100 Then
            GetDynamicYear = CLng(y)
            Exit Function
        End If
    End If

    Dim pos As Long, folderPath As String, fileName As String
    Dim extractedYear As String, tempYear As Long
    Dim patterns As Variant, pattern As Variant

    On Error Resume Next
    folderPath = ThisWorkbook.Path
    fileName = ThisWorkbook.Name
    On Error GoTo 0

    patterns = Array("Horaire_", "Planning_")

    For Each pattern In patterns
        pos = InStrRev(folderPath, CStr(pattern), -1, vbTextCompare)
        If pos > 0 Then
            extractedYear = Mid(folderPath, pos + Len(CStr(pattern)), 4)
            If IsNumeric(extractedYear) Then
                tempYear = CLng(extractedYear)
                If tempYear > 2000 Then
                    GetDynamicYear = tempYear
                    Exit Function
                End If
            End If
        End If
    Next pattern

    For Each pattern In patterns
        pos = InStrRev(fileName, CStr(pattern), -1, vbTextCompare)
        If pos > 0 Then
            extractedYear = Mid(fileName, pos + Len(CStr(pattern)), 4)
            If IsNumeric(extractedYear) Then
                tempYear = CLng(extractedYear)
                If tempYear > 2000 Then
                    GetDynamicYear = tempYear
                    Exit Function
                End If
            End If
        End If
    Next pattern

    GetDynamicYear = Year(Date)
End Function

Private Function GetLocalWorkbookPath() As String
    ' Convertit le chemin OneDrive (URL) en chemin local si necessaire
    Dim wbPath As String
    Dim oneDriveBase As String
    Dim localBase As String
    
    wbPath = ThisWorkbook.Path
    
    ' Si c'est une URL OneDrive, convertir en chemin local
    If Left(wbPath, 5) = "https" Or Left(wbPath, 4) = "http" Then
        ' Chemin OneDrive personnel
        oneDriveBase = Environ("OneDrive")
        If oneDriveBase <> "" Then
            ' Extraire la partie apres le CID OneDrive
            Dim pos As Long
            pos = InStr(1, wbPath, "/", vbTextCompare)
            If pos > 0 Then
                ' Trouver la partie du chemin apres le domaine
                Dim pathPart As String
                Dim parts() As String
                parts = Split(wbPath, "/")
                
                ' Reconstruire le chemin local
                Dim i As Long
                pathPart = ""
                For i = 4 To UBound(parts)  ' Sauter https://domain/cid/
                    If pathPart = "" Then
                        pathPart = parts(i)
                    Else
                        pathPart = pathPart & "\" & parts(i)
                    End If
                Next i
                
                ' Decoder les caracteres URL
                pathPart = Replace(pathPart, "%20", " ")
                pathPart = Replace(pathPart, "^0", " ")
                
                GetLocalWorkbookPath = oneDriveBase & "\" & pathPart
                Exit Function
            End If
        End If
        
        ' Fallback: demander a l'utilisateur
        GetLocalWorkbookPath = Environ("USERPROFILE") & "\Documents"
    Else
        GetLocalWorkbookPath = wbPath
    End If
End Function

Private Function GetOriginalLineNum(mappedLine As Long, mappedLines() As Long, _
                                    linesArray() As String, offset As Long) As Long
    Dim i As Long
    For i = LBound(mappedLines) To UBound(mappedLines)
        If mappedLines(i) = mappedLine Then
            GetOriginalLineNum = CLng(Trim(linesArray(i)))
            Exit Function
        End If
    Next i
    GetOriginalLineNum = mappedLine - offset
End Function

Private Function ConvertShiftCode(code As String) As String
    Select Case UCase(Trim(code))
        Case "C 19": ConvertShiftCode = "7 11:30 15:30 19"
        Case "C 19*": ConvertShiftCode = "7 11:30 15:30 19*"
        Case "C 20": ConvertShiftCode = "8 12 16 20"
        Case "C 20*": ConvertShiftCode = "8 12 16 20*"
        Case Else: ConvertShiftCode = Trim(code)
    End Select
End Function

Private Function CleanSheetName(sheetName As String, wb As Workbook) As String
    Dim tempName As String, suffix As Long, i As Long
    Dim invalidChars As String
    
    tempName = sheetName
    invalidChars = "[];:/\?*'"
    
    For i = 1 To Len(invalidChars)
        tempName = Replace(tempName, Mid$(invalidChars, i, 1), "_")
    Next i
    
    If Len(tempName) > 31 Then tempName = Left$(tempName, 31)
    
    CleanSheetName = tempName
    suffix = 1
    
    Do While SheetExists(CleanSheetName, wb)
        CleanSheetName = Left$(tempName, 31 - Len(CStr(suffix)) - 2) & "_" & suffix
        suffix = suffix + 1
    Loop
End Function

Private Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

Private Function CleanFileName(fileName As String) As String
    Dim invalidChars As String, i As Long
    invalidChars = ":\/*?""<>|()"
    CleanFileName = fileName
    For i = 1 To Len(invalidChars)
        CleanFileName = Replace(CleanFileName, Mid$(invalidChars, i, 1), "_")
    Next i
    Do While InStr(1, CleanFileName, "__") > 0
        CleanFileName = Replace(CleanFileName, "__", "_")
    Loop
End Function

Private Function EnsureFolderExists(fullPath As String) As Boolean
    Dim parts() As String, i As Long, buildPath As String
    
    On Error GoTo FailSafe
    parts = Split(fullPath, "\")
    If UBound(parts) < 0 Then Exit Function
    
    buildPath = parts(0)
    If InStr(buildPath, ":") > 0 Then buildPath = buildPath & "\"
    
    For i = IIf(InStr(parts(0), ":") > 0, 1, 0) To UBound(parts)
        If buildPath = "" Then
            buildPath = parts(i)
        Else
            If Right(buildPath, 1) <> "\" Then buildPath = buildPath & "\"
            buildPath = buildPath & parts(i)
        End If
        If Dir(buildPath, vbDirectory) = "" Then MkDir buildPath
    Next i
    
    EnsureFolderExists = True
    Exit Function

FailSafe:
    EnsureFolderExists = False
End Function
