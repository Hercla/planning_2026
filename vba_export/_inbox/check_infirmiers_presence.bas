Attribute VB_Name = "check_infirmiers_presence"
Option Explicit

'================================================================================
' CHECK_PRESENCE - VERSION SIMPLIFIÉE
' Lit directement les totaux Matin/AM/Soir/Nuit des lignes du planning
'================================================================================

Public Sub Check_Presence_Simplifie()
    ' Lit les totaux DÉJÀ CALCULÉS dans le planning (lignes 10, 11, 12)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If Not TypeOf ws Is Worksheet Then
        MsgBox "Selectionnez une feuille de planning.", vbExclamation
        Exit Sub
    End If
    
    ' Config - lignes des totaux
    Dim rowMatin As Long, rowAM As Long, rowSoir As Long, rowNuit As Long
    Dim minInf As Long, headerRow As Long
    
    rowMatin = CLng(CfgValueOr("CHK_RowMatin", 10))
    rowAM = CLng(CfgValueOr("CHK_RowAM", 11))
    rowSoir = CLng(CfgValueOr("CHK_RowSoir", 12))
    rowNuit = CLng(CfgValueOr("CHK_RowNuit", 0)) ' 0 = pas de ligne nuit sur Jour
    headerRow = CLng(CfgValueOr("CHK_HeaderRow", 4))
    minInf = CLng(CfgValueOr("CHK_MinInfJour", 2))
    
    ' Trouver colonnes des jours
    Dim colStart As Long, colEnd As Long, j As Long
    colStart = FindFirstDayCol(ws, headerRow)
    colEnd = FindLastDayCol(ws, headerRow, colStart)
    
    ' Vérifier chaque jour
    Dim issM As String, issA As String, issS As String
    Dim dayNum As Variant, valM As Double, valA As Double, valS As Double
    
    For j = colStart To colEnd
        dayNum = ws.Cells(headerRow, j).Value
        If IsNumeric(dayNum) Then
            valM = Val(ws.Cells(rowMatin, j).Value & "")
            valA = Val(ws.Cells(rowAM, j).Value & "")
            valS = Val(ws.Cells(rowSoir, j).Value & "")
            
            If valM < minInf Then issM = issM & IIf(issM <> "", ",", "") & dayNum
            If valA < minInf Then issA = issA & IIf(issA <> "", ",", "") & dayNum
            If valS < minInf Then issS = issS & IIf(issS <> "", ",", "") & dayNum
        End If
    Next j
    
    ' Résultat
    Dim msg As String
    If issM = "" And issA = "" And issS = "" Then
        msg = "OK: Min " & minInf & " INF atteint tous les jours!"
    Else
        msg = "Effectif INF insuffisant (min " & minInf & "):" & vbCrLf
        If issM <> "" Then msg = msg & "Matin: " & issM & vbCrLf
        If issA <> "" Then msg = msg & "Après-midi: " & issA & vbCrLf
        If issS <> "" Then msg = msg & "Soir: " & issS
    End If
    MsgBox msg, vbInformation, "Resultat - Présence INF"
End Sub

'================================================================================
' ANCIENNE VERSION - Check_Presence_Infirmiers (calcul détaillé)
'================================================================================

Public Sub Check_Presence_Infirmiers()
    Dim team As String
    team = Trim(InputBox("Equipe a verifier (Jour ou Nuit):", "Verification", "Jour"))
    If team = "" Then Exit Sub
    team = UCase(team)
    
    If team <> "JOUR" And team <> "NUIT" Then
        MsgBox "Saisir 'Jour' ou 'Nuit'.", vbCritical
        Exit Sub
    End If
    
    If Not TypeOf ActiveSheet Is Worksheet Then
        MsgBox "Selectionnez une feuille de planning.", vbExclamation
        Exit Sub
    End If
    
    CheckPresenceForTeam ActiveSheet, team
End Sub

'================================================================================
' CORE - Uses Position from Personnel sheet
'================================================================================

Private Sub CheckPresenceForTeam(ByVal ws As Worksheet, ByVal teamName As String)
    Dim headerRow As Long, minInfJour As Long, minInfNuit As Long
    Dim infFunctions As Variant
    
    ' Load config
    headerRow = CLng(CfgValueOr("CHK_HeaderRow", 4))
    minInfJour = CLng(CfgValueOr("CHK_MinInfJour", 2))
    minInfNuit = CLng(CfgValueOr("CHK_MinInfNuit", 1))
    infFunctions = Split(CfgTextOr("CHK_InfFunctions", "INF;IC"), ";")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUp
    
    ' Find day columns in planning
    Dim colStart As Long, lastCol As Long, dayCount As Long, j As Long
    colStart = FindFirstDayCol(ws, headerRow)
    lastCol = FindLastDayCol(ws, headerRow, colStart)
    dayCount = lastCol - colStart + 1
    
    ' Get month name from planning sheet name
    Dim monthName As String
    monthName = GetMonthFromSheetName(ws.Name)
    If monthName = "" Then
        MsgBox "Impossible de determiner le mois depuis le nom de l'onglet: " & ws.Name, vbCritical
        GoTo CleanUp
    End If
    
    ' Find Personnel sheet
    Dim wsPers As Worksheet
    On Error Resume Next
    Set wsPers = ThisWorkbook.Sheets(CfgTextOr("SHEET_Personnel", "Personnel"))
    On Error GoTo CleanUp
    If wsPers Is Nothing Then
        MsgBox "Onglet Personnel introuvable.", vbCritical
        GoTo CleanUp
    End If
    
    ' Find position column for this month in Personnel
    Dim posCol As Long, funcCol As Long, equipeCol As Long
    Dim h As String, c As Long
    posCol = 0: funcCol = 0: equipeCol = 0
    
    For c = 1 To 60
        h = UCase(Trim$(CStr(wsPers.Cells(1, c).Value & "")))
        h = Replace(h, "É", "E"): h = Replace(h, "È", "E")
        If InStr(1, h, UCase(monthName), vbTextCompare) > 0 And InStr(1, h, "POSITION", vbTextCompare) > 0 Then
            posCol = c: Exit For
        End If
    Next c
    
    ' Also find Fonction and Equipe columns
    For c = 1 To 20
        h = UCase(Trim$(CStr(wsPers.Cells(1, c).Value & "")))
        h = Replace(h, "É", "E"): h = Replace(h, "È", "E")
        If InStr(1, h, "FONCTION", vbTextCompare) > 0 And funcCol = 0 Then funcCol = c
        If InStr(1, h, "EQUIPE", vbTextCompare) > 0 And equipeCol = 0 Then equipeCol = c
    Next c
    
    If posCol = 0 Then
        MsgBox "Colonne '" & monthName & " Position' introuvable dans Personnel.", vbCritical
        GoTo CleanUp
    End If
    If funcCol = 0 Or equipeCol = 0 Then
        MsgBox "Colonnes Fonction ou Equipe introuvables dans Personnel.", vbCritical
        GoTo CleanUp
    End If
    
    ' Load codes dictionary for special codes
    Dim dictCodes As Object
    Set dictCodes = LoadCodesFromConfigCodes()
    
    ' Counters per day
    Dim countM() As Double, countA() As Double, countS() As Double, countN() As Double
    Dim dayNums() As Variant
    ReDim countM(1 To dayCount): ReDim countA(1 To dayCount)
    ReDim countS(1 To dayCount): ReDim countN(1 To dayCount)
    ReDim dayNums(1 To dayCount)
    
    For j = 1 To dayCount
        dayNums(j) = ws.Cells(headerRow, colStart + j - 1).Value
    Next j
    
    ' Loop through Personnel to find INF with positions
    Dim lastRowPers As Long, i As Long
    Dim persFunc As String, persEquipe As String, persPos As Variant
    Dim planRow As Long, codeVal As String, presArr As Variant
    
    lastRowPers = wsPers.Cells(wsPers.Rows.Count, funcCol).End(xlUp).Row
    
    For i = 2 To lastRowPers
        persFunc = Trim$(CStr(wsPers.Cells(i, funcCol).Value & ""))
        persEquipe = UCase(Trim$(CStr(wsPers.Cells(i, equipeCol).Value & "")))
        persPos = wsPers.Cells(i, posCol).Value
        
        ' Check if this person is INF and on the right team
        If Not IsInfFunc(persFunc, infFunctions) Then GoTo NextPerson
        If persEquipe <> teamName Then GoTo NextPerson
        
        ' Get their row in the planning
        If Not IsNumeric(persPos) Then GoTo NextPerson
        planRow = CLng(persPos)
        If planRow < 1 Then GoTo NextPerson
        
        ' Read their codes for each day
        For j = 1 To dayCount
            codeVal = Trim$(CStr(ws.Cells(planRow, colStart + j - 1).Value & ""))
            If codeVal = "" Then GoTo NextDay
            
            ' Calculate presence using proven logic
            presArr = CalculatePresenceFromCode(codeVal)
            
            ' Fallback to Codes_Speciaux if calculation returned zeros
            If presArr(0) = 0 And presArr(1) = 0 And presArr(2) = 0 And presArr(3) = 0 Then
                If dictCodes.Exists(codeVal) Then
                    presArr = dictCodes(codeVal)
                End If
            End If
            
            ' Add to counters
            If teamName = "JOUR" Then
                countM(j) = countM(j) + presArr(0)
                countA(j) = countA(j) + presArr(1)
                countS(j) = countS(j) + presArr(2)
            Else
                countN(j) = countN(j) + presArr(3)
            End If
NextDay:
        Next j
NextPerson:
    Next i
    
    ' Build results
    Dim issM As String, issA As String, issS As String, issN As String
    For j = 1 To dayCount
        If teamName = "JOUR" Then
            If countM(j) < minInfJour Then issM = issM & IIf(issM <> "", ",", "") & dayNums(j)
            If countA(j) < minInfJour Then issA = issA & IIf(issA <> "", ",", "") & dayNums(j)
            If countS(j) < minInfJour Then issS = issS & IIf(issS <> "", ",", "") & dayNums(j)
        Else
            If countN(j) < minInfNuit Then issN = issN & IIf(issN <> "", ",", "") & dayNums(j)
        End If
    Next j
    
    ' Display
    Dim msg As String
    If teamName = "JOUR" Then
        If issM = "" And issA = "" And issS = "" Then
            msg = "OK: Min " & minInfJour & " INF/fraction atteint tous les jours."
        Else
            msg = "Effectif INF insuffisant (min " & minInfJour & "):" & vbCrLf
            If issM <> "" Then msg = msg & "Matin (7h-12h): " & issM & vbCrLf
            If issA <> "" Then msg = msg & "Après-midi (12h-16h30): " & issA & vbCrLf
            If issS <> "" Then msg = msg & "Soir (16h30-20h): " & issS
        End If
    Else
        If issN = "" Then
            msg = "OK: Min " & minInfNuit & " INF/nuit atteint."
        Else
            msg = "Nuits sans INF (min " & minInfNuit & "): " & issN
        End If
    End If
    MsgBox msg, vbInformation, "Resultat - " & teamName

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' Helper to get month name from sheet name (e.g., "Mars" from "Mars 2026" or "Planning Mars")
Private Function GetMonthFromSheetName(ByVal sheetName As String) As String
    Dim months As Variant, m As Variant
    months = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    
    sheetName = UCase(sheetName)
    For Each m In months
        If InStr(1, sheetName, UCase(m), vbTextCompare) > 0 Then
            GetMonthFromSheetName = m
            Exit Function
        End If
    Next m
    GetMonthFromSheetName = ""
End Function

'================================================================================
' HELPER FUNCTIONS
'================================================================================

Private Function NormalizeNameKey(ByVal n As String) As String
    NormalizeNameKey = Replace(Replace(Replace(n, ", ", "_"), ",", "_"), " ", "_")
End Function

Private Function FindFirstDayCol(ws As Worksheet, hRow As Long) As Long
    Dim j As Long
    For j = 1 To 50
        If IsNumeric(ws.Cells(hRow, j).Value) Then FindFirstDayCol = j: Exit Function
    Next j
    FindFirstDayCol = 3
End Function

Private Function FindLastDayCol(ws As Worksheet, hRow As Long, cStart As Long) As Long
    Dim j As Long
    For j = cStart To cStart + 40
        If Not IsNumeric(ws.Cells(hRow, j).Value) Then FindLastDayCol = j - 1: Exit Function
    Next j
    FindLastDayCol = cStart + 30
End Function

Private Function IsInfFunc(ByVal f As String, ByVal funcs As Variant) As Boolean
    Dim i As Long
    f = UCase(f)
    For i = LBound(funcs) To UBound(funcs)
        If InStr(1, f, Trim$(funcs(i)), vbTextCompare) > 0 Then IsInfFunc = True: Exit Function
    Next i
End Function

'================================================================================
' LOAD CODES SPECIAUX - Reads from Codes_Speciaux table (new hybrid approach)
'================================================================================

Private Function LoadCodesFromConfigCodes() As Object
    ' Renamed but kept same function name for compatibility
    ' Now reads from Codes_Speciaux (special codes only)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Dim ws As Worksheet
    
    ' Try new sheet first, then fallback to old
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Codes_Speciaux")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Return empty dictionary, will rely on auto-calculation
        Set LoadCodesFromConfigCodes = d
        Exit Function
    End If
    
    Dim cC As Long, cM As Long, cA As Long, cS As Long, cN As Long
    Dim j As Long, h As String
    
    For j = 1 To 10
        h = UCase(Trim$(CStr(ws.Cells(1, j).Value & "")))
        h = Replace(h, "É", "E"): h = Replace(h, "È", "E")
        If InStr(1, h, "CODE", vbTextCompare) > 0 And cC = 0 Then cC = j
        If InStr(1, h, "MATIN", vbTextCompare) > 0 And cM = 0 Then cM = j
        If InStr(1, h, "AM", vbTextCompare) > 0 And cA = 0 Then cA = j
        If InStr(1, h, "APRES", vbTextCompare) > 0 And cA = 0 Then cA = j
        If InStr(1, h, "SOIR", vbTextCompare) > 0 And cS = 0 Then cS = j
        If InStr(1, h, "NUIT", vbTextCompare) > 0 And cN = 0 Then cN = j
    Next j
    
    ' If Code column not found, return empty dict
    If cC = 0 Then
        Set LoadCodesFromConfigCodes = d
        Exit Function
    End If
    
    Dim lr As Long, i As Long, k As String
    Dim pM As Double, pA As Double, pS As Double, pN As Double
    lr = ws.Cells(ws.Rows.Count, cC).End(xlUp).Row
    
    For i = 2 To lr
        k = Trim$(CStr(ws.Cells(i, cC).Value & ""))
        If k <> "" Then
            pM = 0: pA = 0: pS = 0: pN = 0
            If cM > 0 Then pM = Val(ws.Cells(i, cM).Value & "")
            If cA > 0 Then pA = Val(ws.Cells(i, cA).Value & "")
            If cS > 0 Then pS = Val(ws.Cells(i, cS).Value & "")
            If cN > 0 Then pN = Val(ws.Cells(i, cN).Value & "")
            d(k) = Array(pM, pA, pS, pN)
        End If
    Next i
    Set LoadCodesFromConfigCodes = d
End Function

Private Function NormalizeHeader(ByVal h As Variant) As String
    h = UCase(Trim$(CStr(h & "")))
    ' Handle standard French accent characters
    h = Replace(h, "É", "E"): h = Replace(h, "È", "E"): h = Replace(h, "Ê", "E"): h = Replace(h, "Ë", "E")
    h = Replace(h, "À", "A"): h = Replace(h, "Â", "A"): h = Replace(h, "Ä", "A")
    h = Replace(h, "Ô", "O"): h = Replace(h, "Ö", "O")
    h = Replace(h, "Ù", "U"): h = Replace(h, "Û", "U"): h = Replace(h, "Ü", "U")
    h = Replace(h, "Î", "I"): h = Replace(h, "Ï", "I")
    h = Replace(h, "Ç", "C")
    ' Remove separators
    h = Replace(h, "-", ""): h = Replace(h, "_", ""): h = Replace(h, " ", "")
    NormalizeHeader = h
End Function

'================================================================================
' DEBUG - Diagnostic détaillé pour un jour spécifique
'================================================================================

Public Sub Debug_DayDetail()
    Dim dayNum As Long
    dayNum = Val(InputBox("Numero du jour a analyser (ex: 31):", "Debug", "31"))
    If dayNum < 1 Or dayNum > 31 Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim dictPers As Object, dictCodes As Object
    Set dictPers = LoadPersonnelDict()
    Set dictCodes = LoadCodesFromConfigCodes()
    If dictPers Is Nothing Then MsgBox "Echec Personnel": Exit Sub
    
    Dim infFunctions As Variant
    infFunctions = Split(CfgTextOr("CHK_InfFunctions", "INF;IC"), ";")
    
    Dim headerRow As Long, firstRow As Long, lastRow As Long
    headerRow = CLng(CfgValueOr("CHK_HeaderRow", 4))
    firstRow = CLng(CfgValueOr("CHK_FirstPersonnelRow", 6))
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Find the column for this day
    Dim colStart As Long, targetCol As Long, j As Long
    colStart = FindFirstDayCol(ws, headerRow)
    targetCol = 0
    For j = colStart To colStart + 40
        If ws.Cells(headerRow, j).Value = dayNum Then
            targetCol = j
            Exit For
        End If
    Next j
    
    If targetCol = 0 Then MsgBox "Jour " & dayNum & " non trouve": Exit Sub
    
    Dim msg As String, countM As Long, countA As Long, countS As Long
    msg = "=== JOUR " & dayNum & " (colonne " & targetCol & ") ===" & vbCrLf & vbCrLf
    
    Dim fullName As String, lookupKey As String, nomOnly As String
    Dim persInfo As Variant, persEquipe As String, persFunc As String
    Dim codeVal As String, presArr As Variant, isInf As Boolean
    Dim hStart As Double, hEnd As Double, i As Long
    
    For i = firstRow To lastRow
        fullName = Trim$(CStr(ws.Cells(i, 1).Value & ""))
        If fullName = "" Then GoTo NextDebugRow
        
        lookupKey = NormalizeNameKey(fullName)
        
        ' Try full key, then fallback
        If Not dictPers.Exists(lookupKey) Then
            If InStr(lookupKey, "_") > 0 Then
                nomOnly = Left$(lookupKey, InStr(lookupKey, "_") - 1)
                If Not dictPers.Exists(nomOnly) Then GoTo NextDebugRow
                lookupKey = nomOnly
            Else
                GoTo NextDebugRow
            End If
        End If
        
        persInfo = dictPers(lookupKey)
        persEquipe = persInfo(0): persFunc = persInfo(1)
        If persEquipe <> "JOUR" Then GoTo NextDebugRow
        
        isInf = IsInfFunc(persFunc, infFunctions)
        If Not isInf Then GoTo NextDebugRow
        
        ' This is an INF on JOUR team
        codeVal = Trim$(CStr(ws.Cells(i, targetCol).Value & ""))
        
        ' Calculate presence
        presArr = CalculatePresenceFromCode(codeVal)
        If presArr(0) = 0 And presArr(1) = 0 And presArr(2) = 0 And presArr(3) = 0 Then
            If dictCodes.Exists(codeVal) Then presArr = dictCodes(codeVal)
        End If
        
        msg = msg & fullName & " [" & persFunc & "]" & vbCrLf
        msg = msg & "  Code: '" & codeVal & "'" & vbCrLf
        msg = msg & "  Presence: M=" & presArr(0) & " AM=" & presArr(1) & " S=" & presArr(2) & vbCrLf
        
        If presArr(0) >= 0.5 Then countM = countM + 1
        If presArr(1) >= 0.5 Then countA = countA + 1
        If presArr(2) >= 0.5 Then countS = countS + 1
        
NextDebugRow:
    Next i
    
    msg = msg & vbCrLf & "=== TOTAUX INF ===" & vbCrLf
    msg = msg & "Matin: " & countM & " | AM: " & countA & " | Soir: " & countS
    
    MsgBox msg, vbInformation, "Debug Jour " & dayNum
End Sub


'================================================================================
' LOAD PERSONNEL
'================================================================================

Private Function LoadPersonnelDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(CfgTextOr("SHEET_Personnel", "Personnel")): On Error GoTo 0
    If ws Is Nothing Then MsgBox "Personnel introuvable.", vbCritical: Exit Function
    
    Dim cN As Long, cP As Long, cF As Long, cE As Long, j As Long, h As String, rawH As String
    For j = 1 To 20
        rawH = Trim$(CStr(ws.Cells(1, j).Value & ""))
        h = NormalizeHeader(rawH)
        
        ' Detect NOM column (but not PRENOM)
        If cN = 0 Then
            If h = "NOM" Or (InStr(1, h, "NOM", vbTextCompare) > 0 And InStr(1, h, "PRENOM", vbTextCompare) = 0) Then cN = j
        End If
        
        ' Detect PRENOM column - try multiple patterns
        If cP = 0 Then
            If InStr(1, h, "PRENOM", vbTextCompare) > 0 Then cP = j
            If InStr(1, rawH, "rénom", vbTextCompare) > 0 Then cP = j
            If InStr(1, rawH, "Prénom", vbTextCompare) > 0 Then cP = j
            If InStr(1, rawH, "Prenom", vbTextCompare) > 0 Then cP = j
        End If
        
        If InStr(1, h, "FONCTION", vbTextCompare) > 0 And cF = 0 Then cF = j
        If InStr(1, h, "EQUIPE", vbTextCompare) > 0 And cE = 0 Then cE = j
    Next j
    
    ' Debug: show what columns were found
    ' MsgBox "Nom:" & cN & " Prenom:" & cP & " Fonction:" & cF & " Equipe:" & cE
    
    If cN = 0 Or cF = 0 Or cE = 0 Then
        MsgBox "Colonnes Nom(col " & cN & ")/Fonction(col " & cF & ")/Equipe(col " & cE & ") requises.", vbCritical
        Exit Function
    End If
    
    Dim lr As Long, i As Long, nom As String, prenom As String, fn As String
    Dim equipe As String, fonction As String, infoArr As Variant, nomKey As String
    lr = ws.Cells(ws.Rows.Count, cN).End(xlUp).Row
    
    For i = 2 To lr
        nom = Trim$(CStr(ws.Cells(i, cN).Value))
        prenom = "": If cP > 0 Then prenom = Trim$(CStr(ws.Cells(i, cP).Value))
        
        If nom <> "" Then
            equipe = UCase(Trim$(CStr(ws.Cells(i, cE).Value)))
            fonction = UCase(Trim$(CStr(ws.Cells(i, cF).Value)))
            infoArr = Array(equipe, fonction)
            
            ' Create full key: Nom_Prenom
            fn = Replace(nom, " ", "_")
            If prenom <> "" Then fn = fn & "_" & Replace(prenom, " ", "_")
            
            If Not d.Exists(fn) Then d(fn) = infoArr
            
            ' Also store by nom only as fallback (if not already exists)
            nomKey = Replace(nom, " ", "_")
            If Not d.Exists(nomKey) And prenom <> "" Then d(nomKey) = infoArr
        End If
    Next i
    Set LoadPersonnelDict = d
End Function

'================================================================================
' CALCULATE PRESENCE FROM CODE - Fallback for unknown codes
'================================================================================

Private Function CalculatePresenceFromCode(ByVal code As String) As Variant
    ' Basé sur la logique éprouvée de AutoCategoriserEtColorerHoraires_Final
    Dim pM As Double, pA As Double, pS As Double, pN As Double
    pM = 0: pA = 0: pS = 0: pN = 0
    
    code = Trim$(code)
    If code = "" Then
        CalculatePresenceFromCode = Array(0, 0, 0, 0)
        Exit Function
    End If
    
    Dim codeUpper As String
    codeUpper = UCase(code)
    
    ' Vérification si c'est un code de congé ou absence
    Dim isLeaveCode As Boolean
    isLeaveCode = False
    
    If Left$(codeUpper, 2) = "F " Or Left$(codeUpper, 2) = "R " Then
        isLeaveCode = True
    Else
        Select Case codeUpper
            Case "WE", "ANC", "CA", "CEP", "CP", "CS", "CSS", "CTR", "RCT", "RHS", "RV", "VJ", _
                 "FOR", "MAL", "MAL-MUT", "MAL-GAR", "DP", "EL", "PREAVIS", "PETIT CHOM", _
                 "PAT-EMP", "PAT-MUT", "MAT-EMP", "MAT-MUT", "C SOC", "CRP 1", "DECES", "DEMENAG"
                isLeaveCode = True
        End Select
    End If
    
    If isLeaveCode Then
        CalculatePresenceFromCode = Array(0, 0, 0, 0)
        Exit Function
    End If
    
    ' Extraire les heures du code
    Dim heures As Variant
    heures = ExtraireHeuresSimple(code)
    
    If Not IsArray(heures) Then
        CalculatePresenceFromCode = Array(0, 0, 0, 0)
        Exit Function
    End If
    
    ' Calculer les présences basé sur les fenêtres
    Dim j As Long, hDeb As Double, hFin As Double, overlap As Double
    
    For j = LBound(heures) To UBound(heures) - 1 Step 2
        hDeb = heures(j)
        hFin = heures(j + 1)
        If hFin <= hDeb Then hFin = hFin + 24
        
        ' Matin (Fenêtre 7h-12h)
        overlap = Application.Max(0, Application.Min(hFin, 12) - Application.Max(hDeb, 7))
        If hDeb <= 8 And hFin >= 12 Then
            pM = 1
        ElseIf overlap >= 2 Then
            pM = Application.Max(pM, 0.5)
        End If
        
        ' Après-midi (Fenêtre 12h-17h)
        overlap = Application.Max(0, Application.Min(hFin, 17) - Application.Max(hDeb, 12))
        If hDeb <= 13 And hFin >= 16.5 Then
            pA = 1
        ElseIf overlap >= 2 Then
            pA = Application.Max(pA, 0.5)
        End If
        
        ' Soir (Fenêtre 17h-20h15)
        overlap = Application.Max(0, Application.Min(hFin, 20.25) - Application.Max(hDeb, 17))
        If hDeb < 17.5 And hFin >= 19 Then
            pS = 1
        ElseIf overlap >= 2 Then
            pS = Application.Max(pS, 0.5)
        End If
        
        ' Nuit
        If hDeb >= 20 Or hFin > 24 Then
            pN = 1
        End If
    Next j
    
    ' FORCER les valeurs pour les codes C spéciaux
    Select Case codeUpper
        Case "C 19", "C 19 SA", "C 19 DI", "C 20", "C 20 E", "C 15", "C 15 SA", "C 15 DI"
            pM = 1: pA = 0: pS = 1
        Case "7 15:30", "7 15 30"
            pM = 1: pA = 1: pS = 0
        Case "6:45 15:15", "6 45 15 15"
            pM = 1: pA = 1: pS = 0
    End Select
    
    ' Si poste de Nuit, forcer Soir à 0
    If pN = 1 Then
        Select Case codeUpper
            Case "19:45 6:45", "20 7", "20 24"
                pS = 0
        End Select
    End If
    
    CalculatePresenceFromCode = Array(pM, pA, pS, pN)
End Function

' Fonction simplifiée d'extraction des heures (basée sur ExtraireHeures éprouvée)
Private Function ExtraireHeuresSimple(ByVal code As String) As Variant
    On Error GoTo ErrHandler
    
    code = Replace(code, "-", " ")
    code = Trim$(code)
    If code = "" Then GoTo ErrHandler
    
    Dim parts() As String, cleanParts() As String
    Dim numClean As Long, i As Long
    Dim p As Variant
    
    parts = Split(code, " ")
    ReDim cleanParts(0 To UBound(parts))
    numClean = 0
    
    For Each p In parts
        If Len(CStr(p)) > 0 Then
            If IsNumeric(Left$(CStr(p), 1)) Then
                cleanParts(numClean) = CStr(p)
                numClean = numClean + 1
            End If
        End If
    Next p
    
    If numClean = 0 Or numClean Mod 2 <> 0 Then GoTo ErrHandler
    
    ReDim Preserve cleanParts(0 To numClean - 1)
    
    Dim result() As Double
    ReDim result(1 To numClean)
    
    For i = 0 To numClean - 1
        result(i + 1) = ConvertTimeToDecimalSimple(cleanParts(i))
    Next i
    
    ExtraireHeuresSimple = result
    Exit Function
    
ErrHandler:
    ExtraireHeuresSimple = False
End Function

Private Function ConvertTimeToDecimalSimple(ByVal timeString As String) As Double
    Dim cleanString As String, i As Long, char As String
    
    cleanString = ""
    For i = 1 To Len(timeString)
        char = Mid$(timeString, i, 1)
        If IsNumeric(char) Or char = ":" Or char = "." Or char = "," Then
            cleanString = cleanString & char
        Else
            Exit For
        End If
    Next i
    
    cleanString = Replace(cleanString, ",", ".")
    
    If InStr(cleanString, ":") > 0 Then
        Dim timeParts() As String
        timeParts = Split(cleanString, ":")
        ConvertTimeToDecimalSimple = Val(timeParts(0)) + Val(timeParts(1)) / 60
    Else
        ConvertTimeToDecimalSimple = Val(cleanString)
    End If
End Function

'================================================================================
' EXTRACT HOURS FROM CODE - Handles various formats
'================================================================================

Private Function ExtractHoursFromCode(ByVal code As String, ByRef hStart As Double, ByRef hEnd As Double) As Boolean
    ExtractHoursFromCode = False
    hStart = 0: hEnd = 0
    
    code = Trim$(code)
    If code = "" Then Exit Function
    
    ' Check for special codes in config (SPECIAL_C19, SPECIAL_C20, etc.)
    Dim specialKey As String, specialValue As String
    specialKey = "SPECIAL_" & Replace(UCase(code), " ", "")
    specialValue = CfgTextOr(specialKey, "")
    
    If specialValue <> "" Then
        ' Split shift found - parse all hours and use first/last
        ExtractHoursFromCode = ExtractFirstLastHours(specialValue, hStart, hEnd)
        Exit Function
    End If
    
    ' Handle C XX codes without special config - use simple estimation
    If Left$(UCase(code), 2) = "C " Then
        Dim cParts() As String, cHour As Double
        cParts = Split(code, " ")
        If UBound(cParts) >= 1 Then
            If IsHourLike(cParts(1)) Then
                On Error Resume Next
                cHour = ParseHour(cParts(1))
                On Error GoTo 0
                If cHour >= 12 And cHour <= 24 Then
                    hStart = 7  ' Default start for C codes
                    hEnd = cHour
                    ExtractHoursFromCode = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Split by space
    Dim parts() As String, i As Long
    Dim firstHour As String, lastHour As String
    Dim foundFirst As Boolean: foundFirst = False
    
    parts = Split(code, " ")
    
    ' Find FIRST and LAST numeric parts (hours)
    For i = LBound(parts) To UBound(parts)
        If IsHourLike(parts(i)) Then
            If Not foundFirst Then
                firstHour = parts(i)
                foundFirst = True
            End If
            lastHour = parts(i)  ' Keep updating to get the last one
        End If
    Next i
    
    ' Need both first and last hour values
    If Not foundFirst Then Exit Function
    If firstHour = lastHour Then Exit Function  ' Only one hour found
    
    On Error GoTo ParseError
    hStart = ParseHour(firstHour)
    hEnd = ParseHour(lastHour)
    On Error GoTo 0
    
    ' Validate reasonable values
    If hStart < 0 Or hStart > 24 Then Exit Function
    If hEnd < 0 Or hEnd > 24 Then Exit Function
    
    ExtractHoursFromCode = True
    Exit Function
    
ParseError:
    ExtractHoursFromCode = False
End Function

Private Function ExtractFirstLastHours(ByVal timeStr As String, ByRef hStart As Double, ByRef hEnd As Double) As Boolean
    ' Parse a time string like "7:00 11:30 15:30 19:00" and extract first/last hours
    ExtractFirstLastHours = False
    hStart = 0: hEnd = 0
    
    timeStr = Trim$(timeStr)
    If timeStr = "" Then Exit Function
    
    Dim parts() As String, i As Long
    Dim firstHour As String, lastHour As String
    Dim foundFirst As Boolean: foundFirst = False
    
    parts = Split(timeStr, " ")
    
    For i = LBound(parts) To UBound(parts)
        If IsHourLike(parts(i)) Then
            If Not foundFirst Then
                firstHour = parts(i)
                foundFirst = True
            End If
            lastHour = parts(i)
        End If
    Next i
    
    If Not foundFirst Then Exit Function
    If firstHour = "" Or lastHour = "" Then Exit Function
    
    On Error GoTo ParseErr
    hStart = ParseHour(firstHour)
    hEnd = ParseHour(lastHour)
    On Error GoTo 0
    
    ExtractFirstLastHours = True
    Exit Function
    
ParseErr:
    ExtractFirstLastHours = False
End Function

Private Function IsHourLike(ByVal s As String) As Boolean
    ' Check if string looks like an hour (number or number:number)
    s = Trim$(s)
    If s = "" Then IsHourLike = False: Exit Function
    
    Dim i As Long, c As String, hasDigit As Boolean
    hasDigit = False
    
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then
            hasDigit = True
        ElseIf c <> ":" And c <> "." And c <> "," Then
            IsHourLike = False
            Exit Function
        End If
    Next i
    
    IsHourLike = hasDigit
End Function

Private Function ParseHour(ByVal s As String) As Double
    s = Trim$(s)
    s = Replace(s, ",", ".")
    
    If InStr(s, ":") > 0 Then
        Dim p() As String: p = Split(s, ":")
        ParseHour = Val(p(0)) + Val(p(1)) / 60
    Else
        ParseHour = Val(s)
    End If
End Function

Private Function CalcOverlap(ByVal s1 As Double, ByVal e1 As Double, ByVal s2 As Double, ByVal e2 As Double) As Double
    CalcOverlap = Application.Max(0, Application.Min(e1, e2) - Application.Max(s1, s2))
End Function
