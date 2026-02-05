Attribute VB_Name = "ModuleVerificationCTR"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

' ========================================================================================
' MACRO V7 : VERIFICATION CTR (BASÉE SUR L'ONGLET CONFIG_CODES)
' ========================================================================================

Sub CTR_CheckWeekendEligibility()
    On Error GoTo ErrorHandler

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsCurrent As Worksheet, wsPrev As Worksheet, wsConfig As Worksheet
    Dim wbPrev As Workbook
    Dim baseName As String, shiftType As String
    Dim currentYear As Integer, monthDate As Date, prevMonthDate As Date
    Dim startRow As Long, lastRow As Long, headerRow As Long
    Dim startCol As Long, endCol As Long, configCol As Long
    
    Dim header As Variant, sanitizedHeader() As String
    Dim currentNames As Variant, prevData As Variant
    Dim prevNameIndex As Object
    
    Dim row As Long, j As Long, prevRowIdx As Long
    Dim employeeName As String, employeesWithoutWeekend As String
    Dim validShifts As Object

    Set wsCurrent = ActiveSheet
    
    ' 1. Année
    currentYear = GetPlanningYear()
    
    ' 2. Nom Onglet
    baseName = wsCurrent.Name
    baseName = Replace(baseName, " nuit", "", , , vbTextCompare)
    baseName = Replace(baseName, " jour", "", , , vbTextCompare)
    baseName = Trim(baseName)

    ' 3. Config
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Configuration_CTR_CheckWeek")
    On Error GoTo ErrorHandler
    If wsConfig Is Nothing Then MsgBox "Onglet 'Configuration_CTR_CheckWeek' introuvable.", vbCritical: GoTo CleanUp

    ' 4. Jour/Nuit
    shiftType = DetectShiftType(wsCurrent, wsConfig)
    If shiftType = "" Then
        MsgBox "Impossible de déterminer Jour/Nuit.", vbExclamation
        GoTo CleanUp
    End If

    ' 5. Grille
    configCol = IIf(shiftType = "jour", 2, 3)
    startRow = wsConfig.Cells(2, configCol).value
    lastRow = wsConfig.Cells(3, configCol).value
    headerRow = wsConfig.Cells(4, configCol).value
    startCol = wsConfig.Cells(5, configCol).value
    endCol = wsConfig.Cells(6, configCol).value
    
    ' --- CHARGEMENT DES CODES VALIDES (NOUVEAU) ---
    ' Charge depuis Configuration_CTR_CheckWeek ET depuis Config_Codes (Travail)
    Set validShifts = LoadAllValidCodes(wsConfig)

    ' 6. Dates
    monthDate = GetSafeDateFromSheetName(baseName, currentYear)
    If monthDate = CDate(0) Then MsgBox "Date introuvable pour : " & baseName, vbCritical: GoTo CleanUp
    prevMonthDate = DateAdd("m", -1, monthDate)
    
    ' 7. Onglet Précédent
    Set wsPrev = ResolvePreviousSheet(monthDate, prevMonthDate, shiftType, wbPrev)
    If wsPrev Is Nothing Then
        MsgBox "? Onglet mois précédent introuvable.", vbCritical
        GoTo CleanUp
    End If

    ' 8. Headers
    header = wsPrev.Range(wsPrev.Cells(headerRow, startCol), wsPrev.Cells(headerRow, endCol)).value
    If Not IsArray(header) Then
        ReDim sanitizedHeader(1 To 1): sanitizedHeader(1) = LCase(Trim(CStr(header)))
    Else
        ReDim sanitizedHeader(1 To UBound(header, 2))
        For j = 1 To UBound(header, 2)
            sanitizedHeader(j) = LCase(Trim(CStr(header(1, j))))
        Next j
    End If

    ' 9. Données
    prevData = wsPrev.Range(wsPrev.Cells(startRow, 1), wsPrev.Cells(lastRow, endCol)).value
    currentNames = wsCurrent.Range(wsCurrent.Cells(startRow, 1), wsCurrent.Cells(lastRow, 1)).value

    ' 10. Indexation
    Set prevNameIndex = CreateObject("Scripting.Dictionary")
    prevNameIndex.CompareMode = vbTextCompare
    Dim tmpName As String
    For row = 1 To UBound(prevData, 1)
        tmpName = Trim(CStr(prevData(row, 1)))
        If tmpName <> "" Then If Not prevNameIndex.Exists(tmpName) Then prevNameIndex.Add tmpName, row
    Next row

    ' 11. Vérification
    employeesWithoutWeekend = ""
    For row = 1 To UBound(currentNames, 1)
        employeeName = Trim(CStr(currentNames(row, 1)))
        If employeeName <> "" Then
            If prevNameIndex.Exists(employeeName) Then
                prevRowIdx = prevNameIndex(employeeName)
                ' Check Week-End
                If Not HasWorkedCompleteWeekendRow(prevData, prevRowIdx, startCol, sanitizedHeader, validShifts) Then
                    employeesWithoutWeekend = employeesWithoutWeekend & employeeName & vbNewLine
                End If
            End If
        End If
    Next row

    ' 12. Rapport
    If employeesWithoutWeekend <> "" Then
        MsgBox "?? Employés sans week-end complet en " & MonthToSheetName(prevMonthDate) & " :" & vbNewLine & vbNewLine & employeesWithoutWeekend, vbExclamation
    Else
        MsgBox "? Tous les employés vérifiés sont éligibles CTR.", vbInformation
    End If

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If Not wbPrev Is Nothing Then wbPrev.Close SaveChanges:=False
    Exit Sub

ErrorHandler:
    MsgBox "Erreur : " & Err.description, vbCritical
    Resume CleanUp
End Sub


' ==============================================================================
' FONCTIONS CLES (CHARGEMENT DES CODES ET NETTOYAGE)
' ==============================================================================

' 1. Chargement COMPLET des codes (Config + Onglet Config_Codes)
Function LoadAllValidCodes(wsConfig As Worksheet) As Object
    Dim d As Object
    Dim c As Range, rng As Range
    Dim wsCodes As Worksheet
    Dim lastLine As Long, k As Long
    Dim codeVal As String, typeVal As String
    
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    ' A. Charger depuis Configuration_CTR_CheckWeek (Colonne E)
    Set rng = wsConfig.Range("E2", wsConfig.Cells(wsConfig.Rows.count, "E").End(xlUp))
    For Each c In rng
        codeVal = CleanCellCode(c.value) ' On nettoie aussi la config
        If codeVal <> "" Then d(codeVal) = 1
    Next c
    
    ' B. Charger depuis Config_Codes (Colonne A = Code, Colonne C = Type)
    On Error Resume Next
    Set wsCodes = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    
    If Not wsCodes Is Nothing Then
        lastLine = wsCodes.Cells(wsCodes.Rows.count, "A").End(xlUp).row
        ' On boucle sur Config_Codes
        For k = 2 To lastLine
            typeVal = Trim(LCase(CStr(wsCodes.Cells(k, 3).value))) ' Col C : Type
            
            ' On prend si le Type est "Travail" (ou similaire)
            If typeVal = "travail" Or typeVal = "poste de travail" Then
                codeVal = CleanCellCode(CStr(wsCodes.Cells(k, 1).value)) ' Col A : Code
                If codeVal <> "" Then
                    d(codeVal) = 1
                End If
            End If
        Next k
    End If
    
    Set LoadAllValidCodes = d
End Function

' 2. Vérification Ligne (Compare Code Nettoyé vs Dictionnaire)
Function HasWorkedCompleteWeekendRow(dataArr As Variant, rIdx As Long, colOffset As Long, headers() As String, shifts As Object) As Boolean
    Dim i As Long, colInArray As Long
    Dim h1 As String, h2 As String
    Dim val1 As String, val2 As String
    
    For i = 1 To UBound(headers) - 1
        h1 = headers(i): h2 = headers(i + 1)
        
        If (Left(h1, 3) = "sam" Or h1 = "sa") And (Left(h2, 3) = "dim" Or h2 = "di") Then
            colInArray = (colOffset - 1) + i
            
            If colInArray + 1 <= UBound(dataArr, 2) Then
                ' NETTOYAGE : Transforme "6:45 [Entrée] 15:15" en "6:45 15:15"
                val1 = CleanCellCode(CStr(dataArr(rIdx, colInArray)))
                val2 = CleanCellCode(CStr(dataArr(rIdx, colInArray + 1)))
                
                ' Comparaison Stricte avec le dictionnaire
                If shifts.Exists(val1) And shifts.Exists(val2) Then
                    HasWorkedCompleteWeekendRow = True: Exit Function
                End If
            End If
        End If
    Next i
    HasWorkedCompleteWeekendRow = False
End Function

' 3. Nettoyage Texte (Rend compatible les codes sur 2 lignes)
Function CleanCellCode(ByVal txt As String) As String
    If txt = "" Then Exit Function
    ' Remplace les sauts de ligne par un espace
    txt = Replace(txt, vbCrLf, " ")
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    ' Supprime les espaces doubles (ex: "8  16:30" devient "8 16:30")
    txt = Application.WorksheetFunction.Trim(txt)
    CleanCellCode = txt
End Function

' --- Fonctions Utilitaires Standards ---

Function MonthToSheetName(d As Date) As String
    Dim m As Variant
    m = Array("", "Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    MonthToSheetName = m(Month(d))
End Function

Function ResolvePreviousSheet(currDate As Date, prevDate As Date, shiftType As String, ByRef wbOut As Workbook) As Worksheet
    Dim fd As FileDialog
    Dim wbTarget As Workbook
    Dim ws As Worksheet, bestMatch As Worksheet
    Dim searchMonth As String, searchShift As String, cleanName As String
    
    searchMonth = LCase(MonthToSheetName(prevDate))
    searchShift = LCase(shiftType)

    If Month(currDate) = 1 Then
        MsgBox "?? Passage d'année. Sélectionnez le planning " & Year(currDate) - 1, vbInformation
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .InitialFileName = ThisWorkbook.path & "\"
            If .Show = -1 Then
                On Error Resume Next
                Set wbOut = Workbooks.Open(.SelectedItems(1), ReadOnly:=True)
                On Error GoTo 0
                If wbOut Is Nothing Then Exit Function
                Set wbTarget = wbOut
            Else
                Exit Function
            End If
        End With
    Else
        Set wbTarget = ThisWorkbook
    End If

    For Each ws In wbTarget.Worksheets
        cleanName = LCase(ws.Name)
        If InStr(cleanName, searchMonth) > 0 Then
            If InStr(cleanName, searchShift) > 0 Then
                Set ResolvePreviousSheet = ws: Exit Function
            End If
            If bestMatch Is Nothing Then Set bestMatch = ws
        End If
    Next ws
    Set ResolvePreviousSheet = bestMatch
End Function

Function GetPlanningYear() As Integer
    On Error Resume Next
    GetPlanningYear = ThisWorkbook.Sheets("Feuil_Config").Range("B2").value
    On Error GoTo 0
    If GetPlanningYear = 0 Then GetPlanningYear = Year(Date)
End Function

Function GetSafeDateFromSheetName(strName As String, y As Integer) As Date
    Dim s As String, m As Integer
    s = LCase(Trim(strName))
    If InStr(s, "jan") > 0 Then m = 1
    If InStr(s, "fev") > 0 Or InStr(s, "fév") > 0 Then m = 2
    If InStr(s, "mar") > 0 Then m = 3
    If InStr(s, "avr") > 0 Then m = 4
    If InStr(s, "mai") > 0 Then m = 5
    If InStr(s, "jui") > 0 And InStr(s, "juil") = 0 Then m = 6
    If InStr(s, "juil") > 0 Then m = 7
    If InStr(s, "aou") > 0 Or InStr(s, "aoû") > 0 Then m = 8
    If InStr(s, "sep") > 0 Then m = 9
    If InStr(s, "oct") > 0 Then m = 10
    If InStr(s, "nov") > 0 Then m = 11
    If InStr(s, "dec") > 0 Or InStr(s, "déc") > 0 Then m = 12
    If m > 0 Then GetSafeDateFromSheetName = DateSerial(y, m, 1)
End Function

Function DetectShiftType(ws As Worksheet, wsConfig As Worksheet) As String
    If ws.Rows(wsConfig.Cells(2, 2).value).Hidden = False Then DetectShiftType = "jour"
    If ws.Rows(wsConfig.Cells(2, 3).value).Hidden = False Then DetectShiftType = "nuit"
    If DetectShiftType = "" Then
        If InStr(1, LCase(ws.Name), "nuit") > 0 Then DetectShiftType = "nuit"
        If InStr(1, LCase(ws.Name), "jour") > 0 Then DetectShiftType = "jour"
    End If
End Function
