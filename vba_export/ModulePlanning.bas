Attribute VB_Name = "ModulePlanning"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

'================================================================================================
' MODULE :          Module_Planning (Ultimate Production Version)
' DESCRIPTION :     V2 + exclusions dynamiques + log + warning UI + blocage seuil + auto-suggestions premium
'================================================================================================

' --- CONSTANTES PLANNING (JOUR) ---
Private Const START_ROW As Long = 6
Private Const END_ROW As Long = 26
Private Const START_COL As Long = 3
Private Const END_COL As Long = 33

' --- CONSTANTES LIGNES TOTAUX ---
Private Const TOTAL_ROW_MATIN As Long = 60
Private Const TOTAL_ROW_APRESMIDI As Long = 61
Private Const TOTAL_ROW_SOIR As Long = 62
Private Const PRESENCE_ROW_P06H45 As Long = 64
Private Const PRESENCE_ROW_P07H8H As Long = 65
Private Const PRESENCE_ROW_P8H1630 As Long = 66
Private Const PRESENCE_ROW_C15 As Long = 67
Private Const PRESENCE_ROW_C20 As Long = 68
Private Const PRESENCE_ROW_C20E As Long = 69
Private Const PRESENCE_ROW_C19 As Long = 70

' --- CONSTANTES NUITS ---
Private Const NIGHT_SHIFT_START_ROW As Long = 31
Private Const NIGHT_SHIFT_END_ROW As Long = 38
Private Const NIGHT_CODE_1 As String = "19:45 6:45"
Private Const NIGHT_CODE_2 As String = "20 7"
Private Const PRESENCE_ROW_NIGHT_1 As Long = 71
Private Const PRESENCE_ROW_NIGHT_2 As Long = 72
Private Const TOTAL_ROW_NUIT As Long = 73

' --- CONFIGURATION PERSONNEL ---
Private Const PERSONNEL_SHEET_NAME As String = "Personnel"
Private Const PERSONNEL_COL_NOM As Long = 2
Private Const PERSONNEL_COL_PRENOM As Long = 3
Private Const PERSONNEL_COL_FONCTION As Long = 5 ' Col E

' --- CONFIG (exclusions & référentiel) ---
Private Const CONFIG_SHEET_EXCLUSIONS As String = "Configuration_CTR_CheckWeek"
Private Const CONFIG_EXCLUSION_HEADER As String = "Statuts_A_Exclure" ' horizontal : K1 puis L1...
Private Const CONFIG_KNOWN_HEADER As String = "Statuts_Connus"        ' horizontal : K2 puis L2...

' --- LOG / ALERTES ---
Private Const CONFIG_LOG_SHEET As String = "Configuration_CTR_CheckWeek"
Private Const CONFIG_LOG_HEADER_CELL As String = "K5"
Private Const CONFIG_LOG_START_CELL As String = "K6"

' --- AUTO-SUGGESTIONS (premium) ---
Private Const CONFIG_SUGGEST_HEADER_CELL As String = "K8"
Private Const CONFIG_SUGGEST_START_CELL As String = "K9"
Private Const CONFIG_SUGGEST_COL As String = "K"
Private Const CONFIG_SUGGEST_HINT_COL As String = "L"
Private Const TYPO_DISTANCE_THRESHOLD As Long = 2 ' <=2 = typo probable

' --- BLOCAGE SI TROP D'INCONNUS ---
Private Const UNKNOWN_STATUS_BLOCK_THRESHOLD As Long = 3 ' X (ajuste)

' --- UI WARNING (optionnel) ---
Private Const USERFORM_NAME As String = "UserForm1"
Private Const BTN_MAJ_FRACTIONS_NAME As String = "btnMajFractions" ' si différent, remplace; sinon laisse

' --- VARIABLES GLOBALES ---
Private ignoreIfYellowOrBlue As Object
Private excludedPeople As Object
Private excludedFuncs As Object
Private knownFuncs As Object
Private codeCache As Object

' --- Runtime state ---
Private gUnknownStatusCount As Long
Private gBlockRun As Boolean

'================================================================================================
'   ALIAS COMPATIBILITÉ (menus/boutons existants)
'================================================================================================
Public Sub UpdateDailyTotals()
    UpdateDailyTotals_V2
End Sub

Public Sub MAJ_Fractions()
    UpdateDailyTotals_V2
End Sub

Public Sub MajFractions()
    UpdateDailyTotals_V2
End Sub

'================================================================================================
'   PROCEDURE PRINCIPALE
'================================================================================================
Public Sub UpdateDailyTotals_V2()
    Dim ws As Worksheet: Set ws = ActiveSheet

    ' Sécurité Onglet
    If ws.Name = PERSONNEL_SHEET_NAME Or InStr(1, ws.Name, "Config", vbTextCompare) > 0 Then
        MsgBox "Stop : Impossible de lancer depuis '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo CleanFail

    ' Init config + exclusions + contrôle
    InitConfigDictionaries
    InitExcludedPeopleAndWarn ' remplit excludedPeople + log/suggestions + gBlockRun
    InitIgnoreDicts
    InitCodeCache

    If gBlockRun Then
        MsgBox "Blocage : " & gUnknownStatusCount & " statut(s) inconnu(s) détecté(s) (seuil = " & UNKNOWN_STATUS_BLOCK_THRESHOLD & ")." & vbCrLf & _
               "Corrige le référentiel '" & CONFIG_KNOWN_HEADER & "' avant de relancer.", vbCritical, "MAJ Fractions bloquée"
        GoTo CleanExit
    End If

    Dim nbRows As Long: nbRows = END_ROW - START_ROW + 1
    Dim nbCols As Long: nbCols = END_COL - START_COL + 1

    Dim schedule As Variant
    schedule = ws.Range(ws.Cells(START_ROW, START_COL), ws.Cells(END_ROW, END_COL)).value

    Dim names As Variant
    names = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(END_ROW, 1)).value

    Dim colIndex As Long, rowIndex As Long
    Dim rawCode As String, cleanCode As String
    Dim personRaw As String, personKey As String
    Dim totals(1 To 10) As Double
    Dim cell As Range
    Dim codeInfo As clsCodeInfo
    Dim nightVals As Variant
    Dim nVal As String
    Dim countNight1 As Double, countNight2 As Double
    Dim k As Long

    Dim targetNight1 As String: targetNight1 = NormalizeString(NIGHT_CODE_1)
    Dim targetNight2 As String: targetNight2 = NormalizeString(NIGHT_CODE_2)

    For colIndex = 1 To nbCols
        Dim i As Long
        For i = 1 To 10: totals(i) = 0: Next i
        countNight1 = 0: countNight2 = 0

        ' ------------------ JOUR ------------------
        For rowIndex = 1 To nbRows
            personRaw = CStr(names(rowIndex, 1))
            personKey = NormalizePersonKey(personRaw)

            If Not excludedPeople.Exists(personKey) Then
                rawCode = CStr(schedule(rowIndex, colIndex))
                cleanCode = NormalizeString(rawCode)

                If cleanCode <> "" Then
                    Set cell = ws.Cells(START_ROW + rowIndex - 1, START_COL + colIndex - 1)

                    If Not ShouldBeIgnored(cell, personKey, cleanCode) Then
                        Set codeInfo = GetCachedCodeInfo(cleanCode)

                        If codeInfo.code <> "INCONNU" Then
                            totals(1) = totals(1) + codeInfo.Fractions(1)
                            totals(2) = totals(2) + codeInfo.Fractions(2)
                            totals(3) = totals(3) + codeInfo.Fractions(3)
                            totals(4) = totals(4) + codeInfo.Fractions(5)
                            totals(5) = totals(5) + codeInfo.Fractions(6)
                            totals(6) = totals(6) + codeInfo.Fractions(7)
                            totals(7) = totals(7) + codeInfo.Fractions(8)
                            totals(8) = totals(8) + codeInfo.Fractions(9)
                            totals(9) = totals(9) + codeInfo.Fractions(10)
                            totals(10) = totals(10) + codeInfo.Fractions(11)
                        End If
                    End If
                End If
            End If
        Next rowIndex

        ' ------------------ NUITS ------------------
        nightVals = ws.Range(ws.Cells(NIGHT_SHIFT_START_ROW, START_COL + colIndex - 1), _
                             ws.Cells(NIGHT_SHIFT_END_ROW, START_COL + colIndex - 1)).value

        For k = 1 To UBound(nightVals, 1)
            nVal = NormalizeString(CStr(nightVals(k, 1)))
            If nVal = targetNight1 Then
                countNight1 = countNight1 + 1
            ElseIf nVal = targetNight2 Then
                countNight2 = countNight2 + 1
            End If
        Next k

        WriteTotalsToSheet ws, START_COL + colIndex - 1, totals, countNight1, countNight2
    Next colIndex

CleanExit:
    Set codeCache = Nothing
    Set excludedPeople = Nothing
    Set excludedFuncs = Nothing
    Set knownFuncs = Nothing
    Set ignoreIfYellowOrBlue = Nothing

    Application.ScreenUpdating = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    ' Toujours restaurer l'état Excel
    Application.ScreenUpdating = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    MsgBox "Erreur : " & Err.Number & " - " & Err.description, vbCritical, "UpdateDailyTotals_V2"
End Sub

'================================================================================================
'   CONFIG : lecture horizontale (Statuts_A_Exclure / Statuts_Connus)
'================================================================================================
Private Sub InitConfigDictionaries()
    Set excludedFuncs = GetHorizontalListFromConfig(CONFIG_EXCLUSION_HEADER)

    ' fallback sécurité
    If excludedFuncs Is Nothing Or excludedFuncs.count = 0 Then
        Set excludedFuncs = CreateObject("Scripting.Dictionary")
        excludedFuncs.CompareMode = vbTextCompare
        excludedFuncs("CFA") = True
        excludedFuncs("CEFA") = True
    End If

    Set knownFuncs = GetHorizontalListFromConfig(CONFIG_KNOWN_HEADER)
End Sub

Private Function GetHorizontalListFromConfig(ByVal headerName As String) As Object
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_EXCLUSIONS)

    Dim h As Range
    Set h = ws.Cells.Find(What:=headerName, LookAt:=xlWhole, LookIn:=xlValues)
    If h Is Nothing Then GoTo Fail

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim c As Long
    c = h.Column + 1
    Do While Trim$(CStr(ws.Cells(h.row, c).value)) <> ""
        d(UCase$(NormalizeString(CStr(ws.Cells(h.row, c).value)))) = True
        c = c + 1
    Loop

    Set GetHorizontalListFromConfig = d
    Exit Function

Fail:
    Set GetHorizontalListFromConfig = Nothing
End Function

'================================================================================================
'   EXCLUSIONS PERSONNES + LOG + AUTO-SUGGESTIONS + UI + BLOCAGE
'================================================================================================
Private Sub InitExcludedPeopleAndWarn()
    Set excludedPeople = CreateObject("Scripting.Dictionary")
    excludedPeople.CompareMode = vbTextCompare

    gUnknownStatusCount = 0
    gBlockRun = False

    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Worksheets(PERSONNEL_SHEET_NAME)

    Dim lastRow As Long
    lastRow = wsP.Cells(wsP.Rows.count, PERSONNEL_COL_NOM).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim arr As Variant
    arr = wsP.Range("B2:E" & lastRow).value ' B=Nom, C=Prenom, E=Fonction (index 4)

    Dim doWarn As Boolean
    doWarn = Not (knownFuncs Is Nothing Or knownFuncs.count = 0)

    Dim unknown As Object
    If doWarn Then
        Set unknown = CreateObject("Scripting.Dictionary")
        unknown.CompareMode = vbTextCompare
    End If

    Dim i As Long, nom As String, prenom As String, func As String, key As String

    For i = 1 To UBound(arr, 1)
        nom = CStr(arr(i, 1))
        prenom = CStr(arr(i, 2))
        func = UCase$(NormalizeString(CStr(arr(i, 4))))

        If doWarn Then
            If func <> "" And Not knownFuncs.Exists(func) Then
                unknown(func) = True
            End If
        End If

        If func <> "" And excludedFuncs.Exists(func) Then
            key = NormalizePersonKey(nom & "_" & prenom)
            excludedPeople(key) = True
            key = NormalizePersonKey(nom & " " & prenom)
            excludedPeople(key) = True
        End If
    Next i

    If doWarn Then
        gUnknownStatusCount = unknown.count

        LogUnknownStatusesAndColor unknown
        UpdateUserFormWarning (gUnknownStatusCount > 0), gUnknownStatusCount

        ' Auto-suggestions premium (sépare normal vs typos probables)
        WriteSuggestedStatuses unknown

        If gUnknownStatusCount > UNKNOWN_STATUS_BLOCK_THRESHOLD Then
            gBlockRun = True
        End If
    Else
        ' Pas de référentiel => nettoie UI/log/suggestions (pas de warning)
        LogUnknownStatusesAndColor Nothing
        UpdateUserFormWarning False, 0
        WriteSuggestedStatuses Nothing
    End If
End Sub

Private Sub LogUnknownStatusesAndColor(ByVal unknown As Object)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_LOG_SHEET)

    ' Nettoyer log (col K)
    ws.Range(CONFIG_LOG_START_CELL & ":" & ws.Cells(ws.Rows.count, "K").Address).ClearContents

    Dim hasUnknown As Boolean
    hasUnknown = Not (unknown Is Nothing) And (unknown.count > 0)

    With ws.Range(CONFIG_LOG_HEADER_CELL)
        If hasUnknown Then
            .Interior.Color = vbRed
            .Font.Color = vbWhite
            .Font.Bold = True
            .value = "LOG_STATUTS_INCONNUS  (?)"
        Else
            .Interior.pattern = xlNone
            .Font.ColorIndex = xlAutomatic
            .Font.Bold = False
            If Trim$(CStr(.value)) = "" Then .value = "LOG_STATUTS_INCONNUS"
            If InStr(1, CStr(.value), "LOG_STATUTS_INCONNUS", vbTextCompare) = 0 Then .value = "LOG_STATUTS_INCONNUS"
        End If
    End With

    If Not hasUnknown Then Exit Sub

    Dim r As Long
    r = ws.Range(CONFIG_LOG_START_CELL).row

    ws.Cells(r, "K").value = "[" & Format(Now, "yyyy-mm-dd hh:nn") & "] Statuts inconnus détectés :"

    Dim k As Variant
    For Each k In unknown.keys
        r = r + 1
        ws.Cells(r, "K").value = " - " & CStr(k)
    Next k
End Sub

Private Sub UpdateUserFormWarning(ByVal showWarn As Boolean, Optional ByVal countWarn As Long = 0)
    On Error Resume Next

    Dim uf As Object
    For Each uf In VBA.userForms
        If uf.Name = USERFORM_NAME Then

            ' OPTION 1 : label "lblWarn" (si tu l'ajoutes)
            If Not uf.Controls("lblWarn") Is Nothing Then
                uf.Controls("lblWarn").Visible = showWarn
                If showWarn Then uf.Controls("lblWarn").Caption = "? " & countWarn
            End If

            ' OPTION 2 : prepend ? au bouton MAJ Fractions si trouvé
            If BTN_MAJ_FRACTIONS_NAME <> "" Then
                If Not uf.Controls(BTN_MAJ_FRACTIONS_NAME) Is Nothing Then
                    Dim baseCaption As String
                    baseCaption = Replace(CStr(uf.Controls(BTN_MAJ_FRACTIONS_NAME).Caption), "?", "")
                    baseCaption = Trim$(baseCaption)
                    If showWarn Then
                        uf.Controls(BTN_MAJ_FRACTIONS_NAME).Caption = "? " & baseCaption
                    Else
                        uf.Controls(BTN_MAJ_FRACTIONS_NAME).Caption = baseCaption
                    End If
                End If
            End If

            Exit For
        End If
    Next uf
End Sub

'================================================================================================
'   AUTO-SUGGESTIONS PREMIUM : sépare "SUGGÉRÉS" vs "TYPOS PROBABLES"
'================================================================================================
Private Sub WriteSuggestedStatuses(ByVal unknown As Object)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_LOG_SHEET)

    ' Clear from K8 down to L (keep other config above intact)
    ws.Range(CONFIG_SUGGEST_HEADER_CELL & ":" & ws.Cells(ws.Rows.count, CONFIG_SUGGEST_HINT_COL).Address).ClearContents

    Dim hasUnknown As Boolean
    hasUnknown = Not (unknown Is Nothing) And (unknown.count > 0)

    If Not hasUnknown Then
        ws.Range(CONFIG_SUGGEST_HEADER_CELL).value = "STATUTS_SUGGÉRÉS (0)"
        ws.Range(CONFIG_SUGGEST_HEADER_CELL).Font.Bold = True
        Exit Sub
    End If

    ' Load keys to array and sort
    Dim n As Long: n = unknown.count
    Dim arr() As String
    ReDim arr(1 To n)

    Dim i As Long: i = 1
    Dim k As Variant
    For Each k In unknown.keys
        arr(i) = CStr(k)
        i = i + 1
    Next k
    SortStringArray arr

    ' Split into normal vs typos
    Dim normalList As Object, typoList As Object, typoHint As Object
    Set normalList = CreateObject("Scripting.Dictionary")
    normalList.CompareMode = vbTextCompare

    Set typoList = CreateObject("Scripting.Dictionary")
    typoList.CompareMode = vbTextCompare

    Set typoHint = CreateObject("Scripting.Dictionary")
    typoHint.CompareMode = vbTextCompare

    Dim doTypoCheck As Boolean
    doTypoCheck = Not (knownFuncs Is Nothing Or knownFuncs.count = 0)

    For i = 1 To n
        Dim s As String
        s = arr(i)

        If doTypoCheck Then
            Dim bestMatch As String, bestDist As Long
            bestMatch = GetClosestKnownStatus(s, bestDist)

            If bestMatch <> "" And bestDist <= TYPO_DISTANCE_THRESHOLD Then
                typoList(s) = True
                typoHint(s) = "Proche de: " & bestMatch & " (d=" & bestDist & ")"
            Else
                normalList(s) = True
            End If
        Else
            normalList(s) = True
        End If
    Next i

    ' Write sections
    Dim r As Long
    r = ws.Range(CONFIG_SUGGEST_HEADER_CELL).row

    ' Section 1 : suggestions
    ws.Cells(r, CONFIG_SUGGEST_COL).value = "STATUTS_SUGGÉRÉS (" & normalList.count & ")  –  " & Format(Now, "yyyy-mm-dd hh:nn")
    With ws.Cells(r, CONFIG_SUGGEST_COL)
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    r = r + 1

    Dim key As Variant
    For Each key In SortedKeys(normalList)
        ws.Cells(r, CONFIG_SUGGEST_COL).value = CStr(key)
        r = r + 1
    Next key

    r = r + 1 ' blank line

    ' Section 2 : typos probables
    ws.Cells(r, CONFIG_SUGGEST_COL).value = "TYPOS PROBABLES (" & typoList.count & ")"
    With ws.Cells(r, CONFIG_SUGGEST_COL)
        .Font.Bold = True
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
    r = r + 1

    For Each key In SortedKeys(typoList)
        With ws.Cells(r, CONFIG_SUGGEST_COL)
            .value = CStr(key)
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
            .Font.Bold = True
        End With
        ws.Cells(r, CONFIG_SUGGEST_HINT_COL).value = typoHint(CStr(key))
        r = r + 1
    Next key

    With ws.Range(CONFIG_SUGGEST_HEADER_CELL & ":" & CONFIG_SUGGEST_HINT_COL & (r - 1))
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With
    ws.Columns(CONFIG_SUGGEST_HINT_COL).AutoFit
End Sub

'================================================================================================
'   TYPO HELPERS : closest known status + Levenshtein
'================================================================================================
Private Function GetClosestKnownStatus(ByVal s As String, ByRef bestDist As Long) As String
    bestDist = 9999
    GetClosestKnownStatus = ""

    If knownFuncs Is Nothing Then Exit Function
    If Len(s) = 0 Then Exit Function

    Dim k As Variant
    For Each k In knownFuncs.keys
        Dim d As Long
        d = LevenshteinDistance(s, CStr(k))
        If d < bestDist Then
            bestDist = d
            GetClosestKnownStatus = CStr(k)
        End If
    Next k
End Function

Private Function LevenshteinDistance(ByVal s As String, ByVal t As String) As Long
    Dim i As Long, j As Long
    Dim n As Long, m As Long
    n = Len(s): m = Len(t)

    If n = 0 Then LevenshteinDistance = m: Exit Function
    If m = 0 Then LevenshteinDistance = n: Exit Function

    Dim d() As Long
    ReDim d(0 To n, 0 To m)

    For i = 0 To n: d(i, 0) = i: Next i
    For j = 0 To m: d(0, j) = j: Next j

    Dim cost As Long
    Dim si As String, tj As String

    For i = 1 To n
        si = Mid$(s, i, 1)
        For j = 1 To m
            tj = Mid$(t, j, 1)
            If si = tj Then cost = 0 Else cost = 1
            d(i, j) = Min3(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = d(n, m)
End Function

Private Function Min3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Min3 = a
    If b < Min3 Then Min3 = b
    If c < Min3 Then Min3 = c
End Function

'================================================================================================
'   SORT HELPERS (arrays + dictionary keys)
'================================================================================================
Private Sub SortStringArray(ByRef arr() As String)
    If (LBound(arr) >= UBound(arr)) Then Exit Sub
    QuickSortStrings arr, LBound(arr), UBound(arr)
End Sub

Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As String

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub

Private Function SortedKeys(ByVal dict As Object) As Variant
    If dict Is Nothing Or dict.count = 0 Then
        SortedKeys = Array()
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(1 To dict.count)

    Dim i As Long: i = 1
    Dim k As Variant
    For Each k In dict.keys
        arr(i) = CStr(k)
        i = i + 1
    Next k

    SortStringArray arr
    SortedKeys = arr
End Function

'================================================================================================
'   SUPPORT : Normalisations
'================================================================================================
Private Function NormalizePersonKey(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    s = Replace(s, Chr(160), " ")
    s = UCase$(Trim$(s))
    s = Replace(s, "-", "_")
    s = Replace(s, " ", "_")
    Do While InStr(s, "__") > 0
        s = Replace(s, "__", "_")
    Loop
    NormalizePersonKey = s
End Function

Private Function NormalizeString(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    s = Replace(s, Chr(160), " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeString = s
End Function

'================================================================================================
'   Cache Codes + Ignore dict
'================================================================================================
Private Sub InitCodeCache()
    Set codeCache = CreateObject("Scripting.Dictionary")
    codeCache.CompareMode = vbTextCompare
End Sub

Private Function GetCachedCodeInfo(ByVal code As String) As clsCodeInfo
    If Not codeCache.Exists(code) Then
        Set codeCache(code) = GetCodeInfo(code)
    End If
    Set GetCachedCodeInfo = codeCache(code)
End Function

Private Sub InitIgnoreDicts()
    Set ignoreIfYellowOrBlue = CreateObject("Scripting.Dictionary")
    ignoreIfYellowOrBlue.CompareMode = vbTextCompare

    ' Exceptions couleur (conserve ton système actuel)
    ignoreIfYellowOrBlue("BOURGEOIS_AURORE|7 15:30") = True
    ignoreIfYellowOrBlue("BOURGEOIS_AURORE|6:45 15:15") = True
    ignoreIfYellowOrBlue("DIALLO_MAMADOU|7 15:30") = True
    ignoreIfYellowOrBlue("DIALLO_MAMADOU|6:45 15:15") = True
    ignoreIfYellowOrBlue("DELA VEGA_EDELYN|7 15:30") = True
    ignoreIfYellowOrBlue("DELA VEGA_EDELYN|6:45 15:15") = True
End Sub

Private Function ShouldBeIgnored(ByVal cell As Range, ByVal normPersonKey As String, ByVal code As String) As Boolean
    Dim key As String: key = normPersonKey & "|" & code
    If Not ignoreIfYellowOrBlue.Exists(key) Then
        ShouldBeIgnored = False
        Exit Function
    End If

    If IsYellow(cell) Or IsLightBlue(cell) Then
        ShouldBeIgnored = True
    Else
        ShouldBeIgnored = False
    End If
End Function

Private Function IsYellow(c As Range) As Boolean
    IsYellow = (c.Interior.Color = vbYellow) Or (c.Interior.ColorIndex = 6)
End Function

Private Function IsLightBlue(c As Range) As Boolean
    On Error Resume Next
    Dim themec As Long: themec = c.Interior.ThemeColor
    Dim tint As Double: tint = c.Interior.TintAndShade
    Dim idx As Long: idx = c.Interior.ColorIndex
    Dim rgbv As Long: rgbv = c.Interior.Color
    IsLightBlue = (themec = xlThemeColorAccent1 And tint > 0) _
                  Or (idx = 37 Or idx = 34 Or idx = 41) _
                  Or (rgbv = RGB(221, 235, 247) Or rgbv = RGB(204, 232, 255) Or rgbv = RGB(198, 239, 255))
End Function

'================================================================================================
'   Ecriture Totaux
'================================================================================================
Private Sub WriteTotalsToSheet(ByVal ws As Worksheet, ByVal col As Long, ByRef totals() As Double, ByVal n1 As Double, ByVal n2 As Double)
    ws.Cells(TOTAL_ROW_MATIN, col).value = IIf(totals(1) > 0, totals(1), "")
    ws.Cells(TOTAL_ROW_APRESMIDI, col).value = IIf(totals(2) > 0, totals(2), "")
    ws.Cells(TOTAL_ROW_SOIR, col).value = IIf(totals(3) > 0, totals(3), "")
    ws.Cells(PRESENCE_ROW_P06H45, col).value = IIf(totals(4) > 0, totals(4), "")
    ws.Cells(PRESENCE_ROW_P07H8H, col).value = IIf(totals(5) > 0, totals(5), "")
    ws.Cells(PRESENCE_ROW_P8H1630, col).value = IIf(totals(6) > 0, totals(6), "")
    ws.Cells(PRESENCE_ROW_C15, col).value = IIf(totals(7) > 0, totals(7), "")
    ws.Cells(PRESENCE_ROW_C20, col).value = IIf(totals(8) > 0, totals(8), "")
    ws.Cells(PRESENCE_ROW_C20E, col).value = IIf(totals(9) > 0, totals(9), "")
    ws.Cells(PRESENCE_ROW_C19, col).value = IIf(totals(10) > 0, totals(10), "")
    ws.Cells(PRESENCE_ROW_NIGHT_1, col).value = IIf(n1 > 0, n1, "")
    ws.Cells(PRESENCE_ROW_NIGHT_2, col).value = IIf(n2 > 0, n2, "")
    ws.Cells(TOTAL_ROW_NUIT, col).value = IIf((n1 + n2) > 0, n1 + n2, "")
End Sub

'================================================================================================
'   Tes macros UI conservées
'================================================================================================
Sub Afficher_cacher_menu()
    If UserForm1.Visible Then
        UserForm1.Hide
        Exit Sub
    End If
    With UserForm1
        .height = 509.25
        .width = 201.75
        .StartUpPosition = 0
        .Left = Application.Left + Application.width - .width - 25
        .Top = Application.Top + 50
        .Show vbModeless
    End With
End Sub

Public Sub InitOngletRoulement()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Roulements")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    With ws
        .Activate
        .Columns("B").ColumnWidth = 20.95
        .Columns("C:BG").ColumnWidth = 4.95
        ActiveWindow.Zoom = 50
        .Cells(1, 1).Select
    End With
    Application.ScreenUpdating = True
End Sub



'====================================================================================
' UI helpers for UserForm buttons (toggles)
'====================================================================================
Public Sub ToggleAsterisqueCellule()
    Dim c As Range
    Set c = ActiveCell
    If c Is Nothing Then Exit Sub
    If Not IsCellInPlanning(c) Then Exit Sub

    Dim v As String
    v = CStr(c.value)
    If Len(v) = 0 Then Exit Sub

    If Right$(v, 1) = "*" Then
        c.value = Left$(v, Len(v) - 1)
    Else
        c.value = v & "*"
    End If
End Sub

Public Sub ToggleColorierCelluleVertFonce()
    ToggleCellColor RGB(0, 176, 80), RGB(255, 255, 255)
End Sub

Public Sub ToggleColorierCelluleBleuClair()
    ToggleCellColor RGB(221, 235, 247), RGB(0, 0, 0)
End Sub

Private Sub ToggleCellColor(ByVal bgColor As Long, ByVal fontColor As Long)
    Dim c As Range
    Set c = ActiveCell
    If c Is Nothing Then Exit Sub
    If Not IsCellInPlanning(c) Then Exit Sub

    If c.Interior.Color = bgColor Then
        c.Interior.Color = vbWhite
        c.Font.Color = vbBlack
    Else
        c.Interior.Color = bgColor
        c.Font.Color = fontColor
    End If
End Sub

Private Function IsCellInPlanning(ByVal c As Range) As Boolean
    On Error Resume Next
    IsCellInPlanning = Not Intersect(c, ActiveSheet.Range("planning")) Is Nothing
    On Error GoTo 0
End Function
