Attribute VB_Name = "Module_Remplacements"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

' ===================================================================================
' MODULE :          Module Principal (Version finale mise à jour)
' DESCRIPTION :     Macro principale + fonctions de support.
'                   - Résolution dynamique de la feuille de configuration (Feuil_Config ou anciens noms)
'                   - Création automatique du classeur de demandes
'                   - Sauvegarde auto avec génération du chemin dynamique
'
' MISE À JOUR :
'   - Zoom 70%
'   - Formatage police (Arial Narrow 16)
'   - Chemin de sauvegarde avec {username} et {annee}
'   - Création récursive du dossier de sauvegarde si manquant
' ===================================================================================

Sub GenerateNewWorkbookAndFillDates_Optimized_V4(nomPrenom As String, dayOrNight As String, postCM As String, ReplacementLines As String)
    'On Error GoTo ErrorHandler ' (décommenter si tu veux un handler "propre")

    ' --- Déclaration des variables ---
    Dim wb As Workbook, newWb As Workbook
    Dim wsSource As Worksheet, wsModel As Worksheet, wsFinal As Worksheet
    Dim savePath As String
    Dim linesArray() As String, mappedLines() As Long
    Dim sourceValues As Variant, cellValue As Variant
    Dim lineIndex As Long, lineNum As Long, sourceLineIdx As Long, j As Long, transposedRow As Long
    Dim yearToUse As Long, monthNumber As Long, tempDate As Date
    Dim monthMapping As Object, dictHolidays As Object
    Dim demandsBySourceLine As Object
    Dim lineDemands As Collection
    Dim demand As CReplacementInfo
    Dim sourceCell As Range
    Dim sheetCreatedCount As Long, originalLineNumber As Long
    Dim daysInMonthFinal As Long, lastDataRow As Long
    Dim hasASData As Boolean, hasIDEData As Boolean
    Dim lineNumKey As Variant
    
    ' --- Feuille de configuration (résolution dynamique) ---
    Dim CONFIG_SHEET_NAME As String
    CONFIG_SHEET_NAME = ResolveConfigSheetName()
    If CONFIG_SHEET_NAME = "" Then
        MsgBox "Erreur: aucune feuille de configuration trouvée ('Feuil_Config' ou anciens noms).", vbCritical
        Exit Sub
    End If

    ' --- Variables pour les paramètres de configuration ---
    Dim lineOffset As Long, asbdColor As Long
    Dim nurseCodes As String, holidayPrefixes As String
    Dim savePathPattern As String
    Dim holidaySheetName As String
    Dim tempVal As Variant

    ' --- Initialisation et Vérifications ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Set wb = ThisWorkbook
    Set wsSource = wb.ActiveSheet

    ' ===============================================================
    ' ÉTAPE 0: CHARGEMENT DES PARAMÈTRES DEPUIS LA FEUILLE CONFIG
    ' ===============================================================
    tempVal = GetConfigValue("DecalageLigneRemplacement", CONFIG_SHEET_NAME)
    If InStr(1, CStr(tempVal), "ERREUR") > 0 Then MsgBox tempVal, vbCritical: GoTo CleanUp
    lineOffset = CLng(tempVal)
    
    tempVal = GetConfigValue("Couleur_ASBD_RGB", CONFIG_SHEET_NAME)
    If InStr(1, CStr(tempVal), "ERREUR") > 0 Then MsgBox tempVal, vbCritical: GoTo CleanUp
    asbdColor = CLng(tempVal)
    
    tempVal = GetConfigValue("CodesInfirmiere", CONFIG_SHEET_NAME)
    If InStr(1, CStr(tempVal), "ERREUR") > 0 Then MsgBox tempVal, vbCritical: GoTo CleanUp
    nurseCodes = CStr(tempVal)
    
    tempVal = GetConfigValue("Prefixe_JourFerie", CONFIG_SHEET_NAME)
    If InStr(1, CStr(tempVal), "ERREUR") > 0 Then MsgBox tempVal, vbCritical: GoTo CleanUp
    holidayPrefixes = CStr(tempVal)

    tempVal = GetConfigValue("CheminSauvegarde", CONFIG_SHEET_NAME)
    If InStr(1, CStr(tempVal), "ERREUR") > 0 Then MsgBox tempVal, vbCritical: GoTo CleanUp
    savePathPattern = CStr(tempVal)
    
    tempVal = GetConfigValue("OngletJoursFeries", CONFIG_SHEET_NAME)
    If InStr(1, CStr(tempVal), "ERREUR") > 0 Then MsgBox tempVal, vbCritical: GoTo CleanUp
    holidaySheetName = CStr(tempVal)
    
    ' --- Déterminer l'année de travail ---
    yearToUse = GetDynamicYear()
    If yearToUse = 0 Then
        MsgBox "Année de travail non déterminée.", vbCritical
        GoTo CleanUp
    End If

    ' --- Charger les jours fériés ---
    Set dictHolidays = GetHolidays(yearToUse, holidayPrefixes, holidaySheetName)
    If dictHolidays Is Nothing Then GoTo CleanUp

    ' --- Feuille modèle ---
    On Error Resume Next
    Set wsModel = wb.Sheets("Model")
    On Error GoTo ErrorHandler
    If wsModel Is Nothing Then
        MsgBox "Feuille 'Model' introuvable!", vbCritical
        GoTo CleanUp
    End If

    ' --- Déterminer le mois à partir de l'onglet actif ---
    Set monthMapping = CreateObject("Scripting.Dictionary")
    InitializeMonthMapping monthMapping
    If Not monthMapping.Exists(LCase(wsSource.Name)) Then
        MsgBox "Onglet '" & wsSource.Name & "' non valide pour un mois.", vbExclamation
        GoTo CleanUp
    End If
    monthNumber = monthMapping(LCase(wsSource.Name))

    ' --- Vérification des lignes demandées ---
    If Trim(ReplacementLines) = "" Then
        MsgBox "Aucune ligne de remplacement spécifiée.", vbExclamation
        GoTo CleanUp
    End If

    linesArray = Split(ReplacementLines, ",")
    ReDim mappedLines(LBound(linesArray) To UBound(linesArray))

    For lineIndex = LBound(linesArray) To UBound(linesArray)
        If IsNumeric(Trim(linesArray(lineIndex))) Then
            mappedLines(lineIndex) = MapReplacementNumber_FromConfig(CLng(Trim(linesArray(lineIndex))), lineOffset)
        Else
            MsgBox "Ligne non numérique: '" & linesArray(lineIndex) & "'.", vbCritical
            GoTo CleanUp
        End If
    Next lineIndex

    ' ===============================================================
    ' ÉTAPE 0.5: CONSTRUCTION DU CHEMIN DE SAUVEGARDE
    ' ===============================================================
    ' On part de la valeur brute dans la config, ex:
    ' C:\Users\{username}\OneDrive\C.1 Admin & Team\1. Hot Topics\Horaire_{annee}\DemandeRemplacements

    savePath = savePathPattern
    savePath = Replace(savePath, "{annee}", yearToUse)
    savePath = Replace(savePath, "{username}", Environ("USERNAME"))

    ' Création récursive du dossier (y compris Horaire_2025 etc.)
    If Not EnsureFolderExists(savePath) Then
        MsgBox "Impossible de créer ou d'accéder au dossier de sauvegarde :" & vbCrLf & savePath, vbCritical
        GoTo CleanUp
    End If
    
    ' ===============================================================
    ' ÉTAPE 1: COLLECTE DES DEMANDES PAR LIGNE SOURCE
    ' ===============================================================
    Set demandsBySourceLine = CreateObject("Scripting.Dictionary")

    For sourceLineIdx = LBound(mappedLines) To UBound(mappedLines)
        lineNum = mappedLines(sourceLineIdx)

        ' On lit les colonnes C -> AG (3 -> 33) pour cette ligne
        sourceValues = wsSource.Range(wsSource.Cells(lineNum, 3), wsSource.Cells(lineNum, 33)).value

        If IsArray(sourceValues) Then
            For j = 1 To UBound(sourceValues, 2)
                cellValue = sourceValues(1, j)

                If Not IsEmpty(cellValue) And Trim(CStr(cellValue)) <> "" Then
                    Set demand = New CReplacementInfo
                    demand.sourceLineNum = lineNum
                    demand.shiftCode = Trim(CStr(cellValue))

                    Set sourceCell = wsSource.Cells(lineNum, j + 2) ' colonne réelle dans la feuille source

                    demand.IsASBD = (sourceCell.Interior.Color = asbdColor)
                    demand.IsNurse = IsNurseShift_FromConfig(demand.shiftCode, nurseCodes)

                    On Error Resume Next
                    tempDate = DateSerial(yearToUse, monthNumber, j)
                    If Err.Number = 0 Then
                        demand.DemandDate = tempDate
                        demand.IsWeekend = (Weekday(demand.DemandDate, vbMonday) >= 6)
                        demand.isHoliday = dictHolidays.Exists(CLng(demand.DemandDate))

                        If Not demandsBySourceLine.Exists(lineNum) Then
                            Set demandsBySourceLine(lineNum) = New Collection
                        End If
                        demandsBySourceLine(lineNum).Add demand
                    End If
                    On Error GoTo ErrorHandler
                End If
            Next j
        End If
    Next sourceLineIdx

    If demandsBySourceLine.count = 0 Then
        MsgBox "Aucune demande valide trouvée pour les lignes spécifiées.", vbInformation
        GoTo CleanUp
    End If

    ' ===============================================================
    ' ÉTAPE 2: CRÉATION DU NOUVEAU CLASSEUR + GÉNÉRATION DES FEUILLES
    ' ===============================================================
    Set newWb = Workbooks.Add(xlWBATWorksheet)

    daysInMonthFinal = Day(DateSerial(yearToUse, monthNumber + 1, 0))
    lastDataRow = 7 + daysInMonthFinal - 1

    Dim d As CReplacementInfo

    For Each lineNumKey In demandsBySourceLine.keys
        Set lineDemands = demandsBySourceLine(lineNumKey)

        originalLineNumber = GetOriginalLineNumber(CLng(lineNumKey), mappedLines, linesArray, lineOffset)

        hasASData = False
        hasIDEData = False

        ' On scanne pour voir si cette ligne a des demandes AS et/ou Infirmières
        For Each d In lineDemands
            If d.IsNurse Then
                hasIDEData = True
            Else
                hasASData = True
            End If
            If hasASData And hasIDEData Then Exit For
        Next d

        ' Feuille AS
        If hasASData Then
            Set wsFinal = CreateAndPrepareSheet(newWb, wsModel, "AS", lineNumKey, originalLineNumber, monthNumber, yearToUse)
            If Not wsFinal Is Nothing Then
                sheetCreatedCount = sheetCreatedCount + 1
                For Each demand In lineDemands
                    If Not demand.IsNurse Then
                        transposedRow = Day(demand.DemandDate) + 7 - 1
                        WriteAndFormatDemandData wsFinal, transposedRow, demand, asbdColor
                    End If
                Next demand
            End If
        End If

        ' Feuille INF
        If hasIDEData Then
            Set wsFinal = CreateAndPrepareSheet(newWb, wsModel, "INF", lineNumKey, originalLineNumber, monthNumber, yearToUse)
            If Not wsFinal Is Nothing Then
                sheetCreatedCount = sheetCreatedCount + 1
                For Each demand In lineDemands
                    If demand.IsNurse Then
                        transposedRow = Day(demand.DemandDate) + 7 - 1
                        WriteAndFormatDemandData wsFinal, transposedRow, demand, asbdColor
                    End If
                Next demand
            End If
        End If
    Next lineNumKey

    ' ===============================================================
    ' ÉTAPE 3: NETTOYAGE DE LA FEUILLE VIDE PAR DÉFAUT
    ' ===============================================================
    Application.DisplayAlerts = False
    If sheetCreatedCount > 0 And newWb.Sheets.count > sheetCreatedCount Then
        On Error Resume Next
        newWb.Sheets(1).Delete
        On Error GoTo ErrorHandler
    ElseIf sheetCreatedCount = 0 Then
        MsgBox "Aucune feuille n'a pu être générée.", vbExclamation
        On Error Resume Next
        newWb.Close SaveChanges:=False
        GoTo CleanUp
    End If
    Application.DisplayAlerts = True

    ' ===============================================================
    ' ÉTAPE 4: FORMAT FINAL + SAUVEGARDE
    ' ===============================================================
    ApplyFinalFormatting newWb, lastDataRow, dictHolidays, asbdColor
    GenerateAndSaveWorkbook newWb, postCM, nomPrenom, dayOrNight, yearToUse, monthNumber, savePath, sheetCreatedCount

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    MsgBox "Erreur #" & Err.Number & ": " & Err.description, vbCritical, "Erreur Inattendue"
    Resume CleanUp
End Sub


'===============================================================================
'   RÉSOLUTIONS DE BASE (CONFIG, JOURS FÉRIÉS, MAPPING, ETC.)
'===============================================================================

Private Function ResolveConfigSheetName() As String
    Dim cand As Variant, nm As Variant
    cand = Array("Feuil_Config", "Configuration_GenerateNewWorkbo", "Configuration_GenerateNewWorkbook")
    For Each nm In cand
        On Error Resume Next
        If Not ThisWorkbook.Sheets(CStr(nm)) Is Nothing Then
            ResolveConfigSheetName = CStr(nm)
            Exit Function
        End If
        On Error GoTo 0
    Next nm
    ResolveConfigSheetName = ""
End Function

Function GetConfigValue(ByVal key As String, ByVal configSheetName As String) As Variant
    Dim wsConfig As Worksheet, configRange As Range
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(configSheetName)
    On Error GoTo 0
    If wsConfig Is Nothing Then
        GetConfigValue = "ERREUR: L'onglet '" & configSheetName & "' est introuvable."
        Exit Function
    End If
    Set configRange = wsConfig.Range("A:A").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    If Not configRange Is Nothing Then
        GetConfigValue = configRange.offset(0, 1).value
    Else
        GetConfigValue = "ERREUR: La clé '" & key & "' est introuvable."
    End If
End Function

Function GetHolidays(ByVal targetYear As Long, ByVal configPrefixes As String, ByVal holidaySheetName As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim wsHolidays As Worksheet
    Dim lastRow&, r&
    Dim holidayData As Variant
    Dim cellValue$, datePart$, dateParts() As String
    Dim dayNum&, monthNum&, holidayDate As Date
    Dim prefixes() As String, prefix As Variant

    On Error Resume Next
    Set wsHolidays = ThisWorkbook.Sheets(holidaySheetName)
    On Error GoTo 0
    If wsHolidays Is Nothing Then
        MsgBox "Erreur critique: L'onglet des jours fériés '" & holidaySheetName & "' spécifié dans la configuration est introuvable.", vbCritical
        Set GetHolidays = Nothing
        Exit Function
    End If

    lastRow = wsHolidays.Cells(wsHolidays.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then
        Set GetHolidays = dict
        Exit Function
    End If

    holidayData = wsHolidays.Range("A2:A" & lastRow).value
    prefixes = Split(configPrefixes, ";")

    If IsArray(holidayData) Then
        For r = LBound(holidayData, 1) To UBound(holidayData, 1)
            cellValue = Trim(CStr(holidayData(r, 1)))
            For Each prefix In prefixes
                If UCase(Left(cellValue, Len(Trim(CStr(prefix))))) = UCase(Trim(CStr(prefix))) Then
                    datePart = Trim(Mid(cellValue, Len(Trim(CStr(prefix))) + 1))
                    dateParts = Split(datePart, "-")
                    If UBound(dateParts) >= 1 Then
                        On Error Resume Next
                        dayNum = CLng(Trim(dateParts(0)))
                        monthNum = CLng(Trim(dateParts(1)))
                        holidayDate = DateSerial(targetYear, monthNum, dayNum)
                        If Err.Number = 0 And Year(holidayDate) = targetYear Then
                            dict(CLng(holidayDate)) = True
                        End If
                        Err.Clear
                        On Error GoTo 0
                    End If
                    Exit For
                End If
            Next prefix
        Next r
    End If

    Set GetHolidays = dict
End Function

Function IsNurseShift_FromConfig(ByVal shiftCode As String, ByVal configString As String) As Boolean
    If Trim(configString) = "" Or Trim(shiftCode) = "" Then Exit Function
    Dim rules() As String, rule As Variant
    rules = Split(configString, ";")
    For Each rule In rules
        If Trim(CStr(rule)) = "*" Then
            If InStr(1, shiftCode, "*", vbTextCompare) > 0 Then
                IsNurseShift_FromConfig = True
                Exit Function
            End If
        Else
            If UCase(Trim(shiftCode)) = UCase(Trim(CStr(rule))) Then
                IsNurseShift_FromConfig = True
                Exit Function
            End If
        End If
    Next rule
End Function

Function MapReplacementNumber_FromConfig(replacementNum As Long, offset As Long) As Long
    MapReplacementNumber_FromConfig = replacementNum + offset
End Function

Function GetOriginalLineNumber(ByVal mappedLine As Long, ByRef mappedLinesArray() As Long, ByRef originalLinesArray() As String, ByVal offset As Long) As Long
    Dim i As Long
    For i = LBound(mappedLinesArray) To UBound(mappedLinesArray)
        If mappedLinesArray(i) = mappedLine Then
            GetOriginalLineNumber = CLng(Trim(originalLinesArray(i)))
            Exit Function
        End If
    Next i
    GetOriginalLineNumber = mappedLine - offset
End Function


'===============================================================================
'   MISE EN FORME FINALE DU CLASSEUR GÉNÉRÉ
'===============================================================================

Sub ApplyFinalFormatting(ByVal targetWb As Workbook, ByVal lastDataRow As Long, ByVal holidays As Object, ByVal colorASBD As Long)
    Dim ws As Worksheet, iRow As Long, dateVal As Variant, targetRange As Range

    For Each ws In targetWb.Worksheets
        ws.Range("A6:F" & lastDataRow).Borders.LineStyle = xlContinuous

        For iRow = 7 To lastDataRow
            If IsDate(ws.Cells(iRow, 2).value) Then
                dateVal = ws.Cells(iRow, 2).value
                Set targetRange = ws.Range("A" & iRow & ":C" & iRow)

                If holidays.Exists(CLng(CDate(dateVal))) Then
                    ' Jour férié = rouge clair
                    With targetRange
                        .Interior.Color = RGB(255, 102, 102)
                        .Font.Bold = True
                        .Font.Color = RGB(0, 0, 0)
                    End With
                ElseIf Weekday(CDate(dateVal), vbMonday) >= 6 Then
                    ' Week-end = jaune clair
                    With targetRange
                        .Interior.Color = RGB(255, 255, 153)
                        .Font.Bold = True
                        .Font.Color = RGB(0, 0, 0)
                    End With
                Else
                    ' Jour normal
                    With targetRange
                        If ws.Cells(iRow, 3).Interior.Color <> colorASBD Then
                            .Interior.ColorIndex = xlNone
                        End If
                        .Font.Bold = False
                        .Font.Color = RGB(0, 0, 0)
                    End With
                End If
            End If
        Next iRow
    Next ws
End Sub

' Écrit chaque demande ligne par ligne dans la feuille finale
Sub WriteAndFormatDemandData(wsTarget As Worksheet, targetRow As Long, demandInfo As CReplacementInfo, ByVal colorASBD As Long)
    Dim cellC As Range
    Set cellC = wsTarget.Cells(targetRow, 3)

    With cellC
        .value = ConvertShift(demandInfo.shiftCode)
        .HorizontalAlignment = xlCenter
        .Font.Name = "Arial Narrow"
        .Font.Size = 16

        ' Les étoiles (*) ne restent que côté IDE pour t'indiquer qu'elle vient de la ligne "infirmière"
        If demandInfo.IsNurse Then
            .value = Replace(.value, "*", "")
        End If

        ' Coloration AS/BD
        If demandInfo.IsASBD Then
            .Interior.Color = colorASBD
        Else
            .Interior.ColorIndex = xlNone
        End If
    End With

    If demandInfo.IsNurse Then
        wsTarget.Cells(targetRow, 5).value = "Infirmière"
    End If
End Sub


'===============================================================================
'   CRÉATION DES FEUILLES ET SAUVEGARDE
'===============================================================================

Sub InitializeMonthMapping(ByRef dict As Object)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    dict.Add "janv", 1: dict.Add "janvier", 1
    dict.Add "fev", 2: dict.Add "février", 2
    dict.Add "mars", 3
    dict.Add "avril", 4
    dict.Add "mai", 5
    dict.Add "juin", 6
    dict.Add "juil", 7: dict.Add "juillet", 7
    dict.Add "aout", 8: dict.Add "août", 8
    dict.Add "sept", 9: dict.Add "septembre", 9
    dict.Add "oct", 10: dict.Add "octobre", 10
    dict.Add "nov", 11: dict.Add "novembre", 11
    dict.Add "dec", 12: dict.Add "décembre", 12
End Sub

Function CreateAndPrepareSheet(ByVal targetWb As Workbook, ByVal modelWs As Worksheet, _
                               ByVal sheetType As String, ByVal sourceLineNum As Long, ByVal originalNum As Long, _
                               ByVal monthNum As Long, ByVal yearNum As Long) As Worksheet
    Dim wsNew As Worksheet
    Dim sheetNameBase As String, sheetNameFinal As String
    Dim titleText As String

    On Error GoTo CreateSheetError

    ' Duplique le modèle
    modelWs.Copy After:=targetWb.Sheets(targetWb.Sheets.count)
    Set wsNew = targetWb.Sheets(targetWb.Sheets.count)

    ' Nom + titre en fonction du type
    If sheetType = "AS" Then
        titleText = "Aide-Soignant (Source L" & sourceLineNum & ")"
        If originalNum <> -1 Then
            sheetNameBase = "AS_Rempl_" & originalNum
        Else
            sheetNameBase = "AS_Ligne_" & sourceLineNum
        End If
    Else
        titleText = "Infirmière (Source L" & sourceLineNum & ")"
        If originalNum <> -1 Then
            Select Case originalNum
                Case 6: sheetNameBase = "INF_Nuit_N°1"
                Case 7: sheetNameBase = "INF_Nuit_N°2"
                Case Else: sheetNameBase = "INF_Rempl_" & originalNum
            End Select
        Else
            sheetNameBase = "INF_Ligne_" & sourceLineNum
        End If
    End If

    sheetNameFinal = CleanSheetName(sheetNameBase, targetWb)
    wsNew.Name = sheetNameFinal
    
    ' Zoom
    wsNew.Activate
    ActiveWindow.Zoom = 70
    
    ' Injecte le titre en F3
    wsNew.Range("F3").value = titleText

    ' Remplit les dates (col A/B à partir de la ligne 7)
    Util_RemplirDatesSub wsNew, monthNum, yearNum

    Set CreateAndPrepareSheet = wsNew
    Exit Function

CreateSheetError:
    MsgBox "Erreur création feuille pour " & sheetType & " Ligne " & sourceLineNum & ":" & vbCrLf & Err.description, vbExclamation
    Set CreateAndPrepareSheet = Nothing
End Function

Private Sub GenerateAndSaveWorkbook(ByVal wbToSave As Workbook, ByVal postCM As String, _
                                  ByVal nomPrenom As String, ByVal dayOrNight As String, _
                                  ByVal yearNum As Long, ByVal monthNum As Long, _
                                  ByVal saveFolderPath As String, ByVal sheetCount As Long)
    Dim baseFileName As String, finalFilename As String

    ' Construction du nom du fichier final
    If UCase(Trim(postCM)) = "/ MOIS" Then
        baseFileName = nomPrenom & "_" & Format(DateSerial(yearNum, monthNum, 1), "mmmm_yyyy")
    ElseIf UCase(Trim(postCM)) = "US 1D" Or UCase(Trim(postCM)) = "US_1D" Then
        baseFileName = "Demande_Remplacement_Us_1D_(" & yearNum & "-" & Format(monthNum, "00") & ")"
    Else
        If Trim(postCM) = "" Then
            baseFileName = nomPrenom & "_" & dayOrNight & "_" & Format(Date, "yyyy-mm-dd")
        Else
            baseFileName = postCM & "_" & nomPrenom & "_" & dayOrNight & "_" & Format(Date, "yyyy-mm-dd")
        End If
    End If

    finalFilename = CleanFileName(baseFileName & ".xlsm")

    On Error Resume Next
    wbToSave.SaveAs fileName:=saveFolderPath & "\" & finalFilename, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    If Err.Number <> 0 Then
        MsgBox "ERREUR SAUVEGARDE :" & vbCrLf & Err.description & vbCrLf & "Fichier: " & saveFolderPath & "\" & finalFilename, vbCritical
        Err.Clear
    Else
        MsgBox "Classeur '" & finalFilename & "' créé avec " & sheetCount & " feuille(s) de demande.", vbInformation
    End If
    On Error GoTo 0
End Sub

Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

Function CleanSheetName(sheetName As String, wb As Workbook) As String
    Dim tempName As String, suffix As Long, invalidChars As String, i As Long
    tempName = sheetName
    invalidChars = "[];:/\?*'"

    For i = 1 To Len(invalidChars)
        tempName = Replace(tempName, Mid$(invalidChars, i, 1), "_")
    Next i

    If Len(tempName) > 31 Then tempName = Left$(tempName, 31)

    CleanSheetName = tempName
    suffix = 1

    Do While SheetExists(CleanSheetName, wb)
        CleanSheetName = Left$(tempName, 31 - (Len(CStr(suffix)) + 2)) & "_(" & suffix & ")"
        suffix = suffix + 1
    Loop
End Function

Sub Util_RemplirDatesSub(ByVal ws As Worksheet, ByVal monthNumber As Long, ByVal dynamicYear As Long)
    Const START_ROW_DEMANDE = 7
    Dim daysInMonth As Long, startDate As Date, i As Long
    Dim datesArr(), daysArr()

    daysInMonth = Day(DateSerial(dynamicYear, monthNumber + 1, 0))
    startDate = DateSerial(dynamicYear, monthNumber, 1)

    ReDim datesArr(1 To daysInMonth, 1 To 1)
    ReDim daysArr(1 To daysInMonth, 1 To 1)

    For i = 1 To daysInMonth
        datesArr(i, 1) = startDate + i - 1
        daysArr(i, 1) = StrConv(Format(startDate + i - 1, "ddd"), vbProperCase)
    Next i

    With ws.Range("A" & START_ROW_DEMANDE).Resize(daysInMonth, 2)
        .ClearContents
        .Columns(1).value = daysArr        ' Jour (lun, mar, ...)
        .Columns(2).value = datesArr       ' Date (dd/mm/yyyy)
        .Columns(2).NumberFormat = "dd/mm/yyyy"
    End With
End Sub

Function ConvertShift(value As String) As String
    Select Case Trim(UCase(value))
        Case "C 19": ConvertShift = "7 11:30 15:30 19"
        Case "C 19*": ConvertShift = "7 11:30 15:30 19*"
        Case "C 20": ConvertShift = "8 12 16 20"
        Case "C 20*": ConvertShift = "8 12 16 20*"
        Case Else: ConvertShift = Trim(value)
    End Select
End Function

Function GetDynamicYear() As Long
    Dim pos As Long
    Dim folderPath$, fileName$, extractedYearStr$
    Dim currentYear&, tempYear&
    Dim patterns As Variant, pattern As Variant

    currentYear = Year(Date)

    On Error Resume Next
    folderPath = ThisWorkbook.path
    fileName = ThisWorkbook.Name
    On Error GoTo 0

    ' Cas normal: on essaie de lire l'année dans le chemin du classeur
    patterns = Array("Horaire_", "Planning_")

    For Each pattern In patterns
        pos = InStrRev(folderPath, CStr(pattern), -1, vbTextCompare)
        If pos > 0 Then
            extractedYearStr = Mid(folderPath, pos + Len(CStr(pattern)), 4)
            If IsNumeric(extractedYearStr) Then
                tempYear = CLng(extractedYearStr)
                If tempYear > 2000 Then
                    GetDynamicYear = tempYear
                    Exit Function
                End If
            End If
        End If
    Next pattern

    ' Sinon: essaie de lire dans le nom du fichier
    For Each pattern In patterns
        pos = InStrRev(fileName, CStr(pattern), -1, vbTextCompare)
        If pos > 0 Then
            extractedYearStr = Mid(fileName, pos + Len(CStr(pattern)), 4)
            If IsNumeric(extractedYearStr) Then
                tempYear = CLng(extractedYearStr)
                If tempYear > 2000 Then
                    GetDynamicYear = tempYear
                    Exit Function
                End If
            End If
        End If
    Next pattern

    ' Dernier recours: année courante
    GetDynamicYear = currentYear
End Function

Function CleanFileName(fileName As String) As String
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


'===============================================================================
'   CRÉATION RÉCURSIVE DU DOSSIER DE SAUVEGARDE
'===============================================================================

Private Function EnsureFolderExists(ByVal fullPath As String) As Boolean
    ' Renvoie True si le dossier final existe (déjà présent ou créé avec succès)
    ' Renvoie False si échec

    Dim parts() As String
    Dim i As Long
    Dim buildPath As String

    On Error GoTo FailSafe

    ' On découpe le chemin absolu en segments séparés par "\"
    parts = Split(fullPath, "\")

    If UBound(parts) < 0 Then
        EnsureFolderExists = False
        Exit Function
    End If

    ' buildPath va se reconstruire morceau par morceau
    ' Exemple fullPath:
    ' C:\Users\hercl\OneDrive\C.1 Admin & Team\1. Hot Topics\Horaire_2025\DemandeRemplacements

    ' Premier segment (souvent "C:"), qu'on remet avec "\"
    buildPath = parts(0)
    If InStr(buildPath, ":") > 0 Then
        buildPath = buildPath & "\"
    End If

    ' On boucle à partir du segment suivant
    For i = IIf(InStr(parts(0), ":") > 0, 1, 0) To UBound(parts)
        If buildPath = "" Then
            buildPath = parts(i)
        Else
            If Right(buildPath, 1) <> "\" Then buildPath = buildPath & "\"
            buildPath = buildPath & parts(i)
        End If

        If Dir(buildPath, vbDirectory) = "" Then
            MkDir buildPath
        End If
    Next i

    EnsureFolderExists = True
    Exit Function

FailSafe:
    EnsureFolderExists = False
End Function




