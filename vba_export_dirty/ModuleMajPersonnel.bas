Attribute VB_Name = "ModuleMajPersonnel"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

' ===========================
'  Version finale auto-mappée + RÉCAP
'  - Corrige le bug "aucune feuille" (matching tolérant)
'  - Détection dynamique des colonnes par en-têtes
'  - Rapport récapitulatif par mois (feuille utilisée + nb écritures)
' ===========================

Private Const PERSONNEL_SHEET_NAME As String = "Personnel"
Private Const NAME_LAST_COL As Long = 2     ' Colonne Nom (B)
Private Const NAME_FIRST_COL As Long = 3    ' Colonne Prénom (C)
Private Const START_DATA_ROW As Long = 2
Private Const START_READ_COL As Long = NAME_LAST_COL
Private Const NUM_MONTHS As Long = 12
Private Const TARGET_WRITE_COL As Long = 1  ' Col A sur les feuilles mois
Private Const MIN_TARGET_ROW As Long = 6

' Ordre des mois utilisé partout
Private monthSheetNames As Variant
Private equivSheetNames As Variant

Sub UpdateMonthlySheets_Final_Polished()
    Dim wsPersonnel As Worksheet
    Dim lastRow As Long
    Dim personnelData As Variant
    Dim dictMonthSheets As Object, dictEquivSheets As Object
    Dim posCols(0 To 11) As Long, pctCols(0 To 11) As Long
    Dim endDataCol As Long
    Dim startTime As Double
    Dim errorCount As Long
    Dim i As Long, m As Long
    
    ' Rappels des étiquettes mois et équivalents numériques
    monthSheetNames = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Sept", "Oct", "Nov", "Dec")
    equivSheetNames = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")

    ' Pour le rapport
    Dim monthSheetUsed(0 To 11) As String, equivSheetUsed(0 To 11) As String
    Dim writesMonth(0 To 11) As Long, writesEquiv(0 To 11) As Long
    For m = 0 To 11
        monthSheetUsed(m) = "(absente)"
        equivSheetUsed(m) = "(absente)"
        writesMonth(m) = 0
        writesEquiv(m) = 0
    Next m

    startTime = Timer
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsPersonnel = ThisWorkbook.Sheets(PERSONNEL_SHEET_NAME)
    lastRow = wsPersonnel.Cells(wsPersonnel.Rows.count, NAME_LAST_COL).End(xlUp).row
    If lastRow < START_DATA_ROW Then GoTo CleanUp

    ' --- 1) Trouver dynamiquement les colonnes "XXX Position" et "XXX %" ---
    If Not MapMonthColumnsByHeaders(wsPersonnel, posCols, pctCols, endDataCol) Then
        MsgBox "Impossible de repérer les colonnes 'Mois Position' / '%' dans l'onglet Personnel.", vbCritical
        GoTo CleanUp
    End If

    ' Charger la zone utile
    personnelData = wsPersonnel.Range(wsPersonnel.Cells(START_DATA_ROW, START_READ_COL), _
                                      wsPersonnel.Cells(lastRow, endDataCol)).Value2
    If Not IsArray(personnelData) Then GoTo CleanUp

    ' --- 2) Feuilles cibles (matching tolérant : accents, suffixes B, indices 1..12) ---
    Set dictMonthSheets = CreateObject("Scripting.Dictionary")
    Set dictEquivSheets = CreateObject("Scripting.Dictionary")
    If Not LoadTargetSheets(monthSheetNames, dictMonthSheets) Then GoTo CleanUp
    If Not LoadTargetSheets(equivSheetNames, dictEquivSheets) Then GoTo CleanUp

    ' Renseigner les noms de feuilles réellement matchées pour le récap
    For m = 0 To 11
        If dictMonthSheets.Exists(monthSheetNames(m)) Then monthSheetUsed(m) = dictMonthSheets(monthSheetNames(m)).Name
        If dictEquivSheets.Exists(equivSheetNames(m)) Then equivSheetUsed(m) = dictEquivSheets(equivSheetNames(m)).Name
    Next m

    ' Index locaux du tableau lu (B..endDataCol)
    Dim lastNameIdx As Long, firstNameIdx As Long
    lastNameIdx = NAME_LAST_COL - START_READ_COL + 1
    firstNameIdx = NAME_FIRST_COL - START_READ_COL + 1

    Dim lastName As String, firstName As String, employeeFullName As String
    Dim targetRow As Variant, pctValue As Variant
    Dim monthlyWs As Worksheet, equivWs As Worksheet

    ' --- 3) Boucle employés ---
    For i = 1 To UBound(personnelData, 1)
        lastName = Trim(CStr(personnelData(i, lastNameIdx)))
        firstName = Trim(CStr(personnelData(i, firstNameIdx)))
        If lastName = "" Or firstName = "" Then GoTo NextEmployee

        employeeFullName = lastName & "_" & firstName

        ' --- Boucle mois ---
        For m = 0 To NUM_MONTHS - 1
            If posCols(m) = 0 Then GoTo NextMonth

            targetRow = personnelData(i, posCols(m) - START_READ_COL + 1)

            ' Valider la ligne demandée
            If IsNumeric(targetRow) And targetRow >= MIN_TARGET_ROW And targetRow = CLng(targetRow) Then
                If pctCols(m) > 0 Then
                    pctValue = personnelData(i, pctCols(m) - START_READ_COL + 1)
                    If IsEmpty(pctValue) Or CStr(pctValue) = "" Then GoTo NextMonth
                End If

                ' Feuilles trouvées
                Set monthlyWs = Nothing
                Set equivWs = Nothing
                If dictMonthSheets.Exists(monthSheetNames(m)) Then Set monthlyWs = dictMonthSheets(monthSheetNames(m))
                If dictEquivSheets.Exists(equivSheetNames(m)) Then Set equivWs = dictEquivSheets(equivSheetNames(m))

                ' Écrire le Nom_Prenom et compter
                If Not monthlyWs Is Nothing Then
                    monthlyWs.Cells(CLng(targetRow), TARGET_WRITE_COL).value = employeeFullName
                    writesMonth(m) = writesMonth(m) + 1
                End If
                If Not equivWs Is Nothing Then
                    equivWs.Cells(CLng(targetRow), TARGET_WRITE_COL).value = employeeFullName
                    writesEquiv(m) = writesEquiv(m) + 1
                End If

            ElseIf CStr(targetRow) <> "" Then
                Debug.Print "Avertissement: Ligne cible invalide (" & CStr(targetRow) & _
                            ") pour " & employeeFullName & " Mois: " & monthSheetNames(m)
                errorCount = errorCount + 1
            End If
NextMonth:
        Next m
NextEmployee:
    Next i

    Dim finishMsg As String
    finishMsg = "Mise à jour terminée en " & Format(Timer - startTime, "0.00") & " s."
    If errorCount > 0 Then finishMsg = finishMsg & vbCrLf & errorCount & " avertissement(s) (voir Ctrl+G)."

    ' --- 4) Bilan lisible ---
    Dim recap As String
    recap = BuildRecap(monthSheetUsed, equivSheetUsed, writesMonth, writesEquiv)
    Debug.Print recap
    MsgBox finishMsg & vbCrLf & vbCrLf & recap, vbInformation, "Récap de l'écriture"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Erreur N°" & Err.Number & " : " & Err.description, vbCritical, "Erreur d'exécution"
    Resume CleanUp
End Sub

' =========================
'     MAPPING PAR EN-TÊTES
' =========================
Private Function MapMonthColumnsByHeaders(ws As Worksheet, _
    ByRef posCols() As Long, ByRef pctCols() As Long, ByRef endCol As Long) As Boolean

    Dim hdrRow As Long: hdrRow = 1
    Dim lastCol As Long
    Dim m As Long
    Dim labelPos As String
    Dim f As Range
    
    lastCol = ws.Cells(hdrRow, ws.Columns.count).End(xlToLeft).Column
    endCol = lastCol
    For m = 0 To 11
        posCols(m) = 0: pctCols(m) = 0
    Next m

    ' Cherche dans la ligne d'en-tête : "<Mois> Position" et "<Mois> %"
    For m = 0 To 11
        labelPos = monthSheetNames(m) & " Position"
        Set f = RowFindCell(ws, hdrRow, labelPos)
        If Not f Is Nothing Then posCols(m) = f.Column
        
        ' Pourcentages : accepte "Mois %", "Mois%" (avec/sans espace)
        Set f = RowFindCell(ws, hdrRow, monthSheetNames(m) & " %")
        If Not f Is Nothing Then pctCols(m) = f.Column
        If pctCols(m) = 0 Then
            Set f = RowFindCell(ws, hdrRow, monthSheetNames(m) & "%")
            If Not f Is Nothing Then pctCols(m) = f.Column
        End If
    Next m

    ' Au moins une colonne Position trouvée ?
    For m = 0 To 11
        If posCols(m) <> 0 Then MapMonthColumnsByHeaders = True: Exit Function
    Next m
    MapMonthColumnsByHeaders = False
End Function

Private Function RowFindCell(ws As Worksheet, hdrRow As Long, ByVal text As String) As Range
    Dim rng As Range
    Set rng = ws.Rows(hdrRow).Find(What:=text, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=False)
    If Not rng Is Nothing Then Set RowFindCell = rng
End Function

' =========================================================
'     MATCHING TOLÉRANT DES FEUILLES (accents/suffixes)
' =========================================================
Private Function NormalizeString(ByVal s As String) As String
    Dim T As String
    T = LCase$(s)
    T = Replace(T, "à", "a"): T = Replace(T, "â", "a"): T = Replace(T, "ä", "a")
    T = Replace(T, "é", "e"): T = Replace(T, "è", "e"): T = Replace(T, "ê", "e"): T = Replace(T, "ë", "e")
    T = Replace(T, "î", "i"): T = Replace(T, "ï", "i")
    T = Replace(T, "ô", "o"): T = Replace(T, "ö", "o")
    T = Replace(T, "ù", "u"): T = Replace(T, "û", "u"): T = Replace(T, "ü", "u")
    T = Replace(T, "ç", "c"): T = Replace(T, "œ", "oe"): T = Replace(T, "'", "")
    T = Replace(T, " ", ""): T = Replace(T, "-", ""): T = Replace(T, "_", "")
    NormalizeString = T
End Function

Private Function FindSheetForKey(ByVal key As String) As Worksheet
    Dim keyN As String, ws As Worksheet, nmN As String
    keyN = NormalizeString(key)
    ' 1) exact normalisé
    For Each ws In ThisWorkbook.Worksheets
        nmN = NormalizeString(ws.Name)
        If nmN = keyN Then Set FindSheetForKey = ws: Exit Function
    Next ws
    ' 2) commence/termine par
    For Each ws In ThisWorkbook.Worksheets
        nmN = NormalizeString(ws.Name)
        If Left$(nmN, Len(keyN)) = keyN Or Right$(nmN, Len(keyN)) = keyN Then
            Set FindSheetForKey = ws: Exit Function
        End If
    Next ws
    ' 3) présence d'un index 1..12
    If IsNumeric(key) Then
        For Each ws In ThisWorkbook.Worksheets
            nmN = NormalizeString(ws.Name)
            If InStr(1, nmN, CStr(CLng(key)), vbTextCompare) > 0 Then
                Set FindSheetForKey = ws: Exit Function
            End If
        Next ws
    End If
End Function

Private Function LoadTargetSheets(sheetNameKeys As Variant, ByRef dictSheets As Object) As Boolean
    Dim key As Variant, ws As Worksheet
    Dim missingKeys As String, firstMissing As Boolean
    
    If dictSheets Is Nothing Then Set dictSheets = CreateObject("Scripting.Dictionary")
    dictSheets.RemoveAll
    dictSheets.CompareMode = vbTextCompare
    
    firstMissing = True
    For Each key In sheetNameKeys
        Set ws = FindSheetForKey(CStr(key))
        If Not ws Is Nothing Then
            If Not dictSheets.Exists(CStr(key)) Then dictSheets.Add CStr(key), ws
        Else
            If Not firstMissing Then missingKeys = missingKeys & ", "
            missingKeys = missingKeys & "'" & CStr(key) & "'"
            firstMissing = False
        End If
        Set ws = Nothing
    Next key
    
    If dictSheets.count = 0 Then
        MsgBox "Aucune des feuilles cibles n'a été trouvée. Opération annulée.", vbCritical, "Erreur"
        LoadTargetSheets = False
    Else
        If Len(missingKeys) > 0 Then Debug.Print "Info: clés non trouvées : " & missingKeys
        LoadTargetSheets = True
    End If
End Function

' =========================================================
'                     RÉCAP FORMATÉ
' =========================================================
Private Function BuildRecap(ByRef monthUsed() As String, ByRef equivUsed() As String, _
                            ByRef wMonth() As Long, ByRef wEquiv() As Long) As String
    Dim m As Long, s As String, totalM As Long, totalE As Long
    s = "Récap par mois :" & vbCrLf
    For m = 0 To 11
        s = s & " - " & monthSheetNames(m) & " ? feuille '" & monthUsed(m) & _
                "', écritures: Mois=" & wMonth(m) & " | Equiv=" & wEquiv(m) & _
                " (clé '" & equivSheetNames(m) & "' ? '" & equivUsed(m) & "')" & vbCrLf
        totalM = totalM + wMonth(m)
        totalE = totalE + wEquiv(m)
    Next m
    s = s & vbCrLf & "TOTAL écritures: Mois=" & totalM & " | Equiv=" & totalE
    BuildRecap = s
End Function


