Attribute VB_Name = "Module_Remplacement_Auto"
'Attribute VB_Name = "Module_Remplacement_Auto"
' ExportedAt: 2026-01-30 | Workbook: Planning_2026.xlsm
' OPTIMISE: Utilise Module_Planning_Core pour fonctions communes
Option Explicit

' =========================================================================
' CONSTANTES LIGNES (Configuration du planning)
' =========================================================================

Const LIGNE_DEBUT_PLANNING_PERSONNEL As Long = 6
Const LIGNE_FIN_PLANNING_PERSONNEL As Long = 30
Const LIGNE_AIDE_SOIGNANT_C19_PLANNING As Long = 24
Const LIGNE_FIN_IDE_PLANNING As Long = 23

Const LIGNE_REMPLACEMENT_DEBUT_JOUR As Long = 40
Const LIGNE_REMPLACEMENT_FIN_JOUR As Long = 41
Const NB_REMPLACEMENT_JOUR_LIGNES As Long = LIGNE_REMPLACEMENT_FIN_JOUR - LIGNE_REMPLACEMENT_DEBUT_JOUR + 1

Const LIGNE_REMPLACEMENT_DEBUT_NUIT As Long = 46
Const LIGNE_REMPLACEMENT_FIN_NUIT As Long = 47
Const NB_REMPLACEMENT_NUIT_LIGNES As Long = LIGNE_REMPLACEMENT_FIN_NUIT - LIGNE_REMPLACEMENT_DEBUT_NUIT + 1

' =========================================================================
' CONSTANTES INDICES CODES
' =========================================================================

Const SUGG_645 As Long = 0
Const SUGG_7_1530 As Long = 1
Const SUGG_7_1130 As Long = 2
Const SUGG_7_13 As Long = 3
Const SUGG_8_1630 As Long = 4
Const SUGG_C15_GRP As Long = 5
Const SUGG_C20_CODE As Long = 6
Const SUGG_C20E_CODE As Long = 7
Const SUGG_C19_CODE As Long = 8
Const SUGG_12_30_16_30 As Long = 9
Const SUGG_NUIT1 As Long = 10
Const SUGG_NUIT2 As Long = 11

' =========================================================================
' FONCTIONS UTILITAIRES LOCALES (Non dans Core car spécifiques)
' =========================================================================

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    If Not IsArray(arr) Then Exit Function
    For i = LBound(arr) To UBound(arr)
        If StrComp(val, CStr(arr(i)), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Function CodeDejaPresent(planningArr As Variant, rempArr As Variant, jourCol As Long, codeToCheck As String, Optional exactMatch As Boolean = False) As Boolean
    Dim r As Long, cellVal As String
    CodeDejaPresent = False
    
    On Error Resume Next
    If Not IsEmpty(planningArr) And IsArray(planningArr) Then
        If jourCol > 0 And jourCol <= UBound(planningArr, 2) And LBound(planningArr, 1) <= UBound(planningArr, 1) Then
            For r = LBound(planningArr, 1) To UBound(planningArr, 1)
                cellVal = Trim(CStr(planningArr(r, jourCol)))
                If exactMatch Then
                    If StrComp(cellVal, codeToCheck, vbTextCompare) = 0 Then CodeDejaPresent = True: Exit Function
                Else
                    If InStr(1, cellVal, codeToCheck, vbTextCompare) > 0 Then CodeDejaPresent = True: Exit Function
                End If
            Next r
        End If
    End If
    
    If Not IsEmpty(rempArr) And IsArray(rempArr) Then
        If jourCol > 0 And jourCol <= UBound(rempArr, 2) And LBound(rempArr, 1) <= UBound(rempArr, 1) Then
            For r = LBound(rempArr, 1) To UBound(rempArr, 1)
                cellVal = Trim(CStr(rempArr(r, jourCol)))
                If exactMatch Then
                    If StrComp(cellVal, codeToCheck, vbTextCompare) = 0 Then CodeDejaPresent = True: Exit Function
                Else
                    If InStr(1, cellVal, codeToCheck, vbTextCompare) > 0 Then CodeDejaPresent = True: Exit Function
                End If
            Next r
        End If
    End If
    On Error GoTo 0
End Function

Function EstUnOngletDeMois(nomFeuille As String) As Boolean
    EstUnOngletDeMois = ( _
        nomFeuille Like "Janv*" Or nomFeuille Like "Fev*" Or nomFeuille Like "Mars*" Or _
        nomFeuille Like "Avril*" Or nomFeuille Like "Mai*" Or nomFeuille Like "Juin*" Or _
        nomFeuille Like "Juil*" Or nomFeuille Like "Aout*" Or nomFeuille Like "Sept*" Or _
        nomFeuille Like "Oct*" Or nomFeuille Like "Nov*" Or nomFeuille Like "Dec*" _
    )
End Function

Function ChoisirCodePertinent(codesPossibles As Variant, planningArr As Variant, rempArr As Variant, col As Long) As String
    Dim code As Variant, freq As Object, r As Long, cellVal As String
    Set freq = CreateObject("Scripting.Dictionary")

    If Not IsArray(codesPossibles) Then
        If IsMissing(codesPossibles) Or IsEmpty(codesPossibles) Or IsNull(codesPossibles) Or codesPossibles = "" Then
            ChoisirCodePertinent = ""
            Exit Function
        Else
            codesPossibles = Array(codesPossibles)
        End If
    ElseIf LBound(codesPossibles) > UBound(codesPossibles) Then
        ChoisirCodePertinent = ""
        Exit Function
    End If
    
    For Each code In codesPossibles
        freq(CStr(code)) = 0
    Next code
    
    If Not IsEmpty(planningArr) And IsArray(planningArr) Then
        If col > 0 And col <= UBound(planningArr, 2) And LBound(planningArr, 1) <= UBound(planningArr, 1) Then
            For r = LBound(planningArr, 1) To UBound(planningArr, 1)
                cellVal = Trim(CStr(planningArr(r, col)))
                If freq.Exists(cellVal) Then freq(cellVal) = freq(cellVal) + 1
            Next r
        End If
    End If
    
    If Not IsEmpty(rempArr) And IsArray(rempArr) Then
        If col > 0 And col <= UBound(rempArr, 2) And LBound(rempArr, 1) <= UBound(rempArr, 1) Then
            For r = LBound(rempArr, 1) To UBound(rempArr, 1)
                cellVal = Trim(CStr(rempArr(r, col)))
                If freq.Exists(cellVal) Then freq(cellVal) = freq(cellVal) + 1
            Next r
        End If
    End If

    For Each code In codesPossibles
        If freq(CStr(code)) = 0 Then ChoisirCodePertinent = CStr(code): Exit Function
    Next code
    
    Dim minFreq As Long: minFreq = 99999
    Dim bestCode As String: bestCode = ""
    If LBound(codesPossibles) <= UBound(codesPossibles) Then
        bestCode = CStr(codesPossibles(LBound(codesPossibles)))
        minFreq = freq(bestCode)
    Else
        ChoisirCodePertinent = ""
        Exit Function
    End If

    For Each code In codesPossibles
        If freq(CStr(code)) < minFreq Then
            minFreq = freq(CStr(code))
            bestCode = CStr(code)
        End If
    Next code
    ChoisirCodePertinent = bestCode
End Function

Sub MettreAJourCompteursMAS(codePlace As String, ByRef actualMatin As Long, ByRef actualPM As Long, ByRef actualSoir As Long, ByVal ligneEnCours As Long, ByRef presence7_8h_compteur As Long)
    Select Case codePlace
        Case "6:45 15:15", "7 15:30", "7 13", "7 11:30"
            actualMatin = actualMatin + 1
            If codePlace = "6:45 15:15" Or codePlace = "7 15:30" Then actualPM = actualPM + 1
            presence7_8h_compteur = presence7_8h_compteur + 1
        Case "8 16:30"
            actualMatin = actualMatin + 1
            actualPM = actualPM + 1
        Case "C 15", "C 15 bis", "C 15 di"
            actualPM = actualPM + 1
            actualSoir = actualSoir + 1
        Case "C 20", "C 20 E"
            actualPM = actualPM + 1
            actualSoir = actualSoir + 1
        Case "C 19", "C 19 di"
            actualMatin = actualMatin + 1
            actualSoir = actualSoir + 1
            presence7_8h_compteur = presence7_8h_compteur + 1
        Case "12:30 16:30"
            actualPM = actualPM + 1
        Case "8:30 12:45 16:30 20:15"
            If ligneEnCours = LIGNE_DEBUT_PLANNING_PERSONNEL Then
                actualPM = actualPM + 1
                actualSoir = actualSoir + 1
            End If
    End Select
End Sub

Function CreateTargetArray(data As Variant) As Variant
    Dim numDays As Long, numConditions As Long
    Dim i As Long, j As Long
    Dim tempArr As Variant

    If Not IsArray(data) Then Exit Function
    If LBound(data) > UBound(data) Then Exit Function

    numDays = UBound(data) - LBound(data) + 1

    If Not IsArray(data(LBound(data))) Then Exit Function
    If LBound(data(LBound(data))) > UBound(data(LBound(data))) Then Exit Function

    numConditions = UBound(data(LBound(data))) - LBound(data(LBound(data))) + 1
    
    ReDim tempArr(0 To numDays - 1, 0 To numConditions - 1) As Long

    For i = 0 To numDays - 1
        If IsArray(data(LBound(data) + i)) And _
           LBound(data(LBound(data) + i)) <= UBound(data(LBound(data) + i)) Then
            For j = 0 To numConditions - 1
                If LBound(data(LBound(data) + i)) + j <= UBound(data(LBound(data) + i)) Then
                    On Error Resume Next
                    tempArr(i, j) = CLng(data(LBound(data) + i)(LBound(data(LBound(data))) + j))
                    If Err.Number <> 0 Then tempArr(i, j) = 0
                    On Error GoTo 0
                Else
                    tempArr(i, j) = 0
                End If
            Next j
        Else
            For j = 0 To numConditions - 1
                tempArr(i, j) = 0
            Next j
        End If
    Next i
    CreateTargetArray = tempArr
End Function

Private Sub ActualiserManquesValeurs(ByRef manqueMatin As Long, ByRef manquePM As Long, ByRef manqueSoir As Long, _
                                   ByVal targetMatin As Long, ByVal targetPM As Long, ByVal targetSoir As Long, _
                                   ByVal actualMatin As Long, ByVal actualPM As Long, ByVal actualSoir As Long)
    manqueMatin = targetMatin - actualMatin: If manqueMatin < 0 Then manqueMatin = 0
    manquePM = targetPM - actualPM: If manquePM < 0 Then manquePM = 0
    manqueSoir = targetSoir - actualSoir: If manqueSoir < 0 Then manqueSoir = 0
End Sub

' =========================================================================
' FONCTION PRINCIPALE : TRAITER UNE FEUILLE DE MOIS
' OPTIMISATION CLÉ : Lit les effectifs depuis lignes 60-62 (Module_Planning_Core)
' =========================================================================

Sub TraiterUneFeuilleDeMois(ws As Worksheet, _
                            LdebFractions As Long, LfinFractions As Long, _
                            colDeb As Long, _
                            groupesExclusifs As Variant, _
                            codesSuggestion As Variant)
    Dim col As Long, jourSemaine As Long, i As Long, l As Long
    Dim dateJour As Date, codeFerie As Boolean
    Dim nbJours As Long
    Dim planningArr As Variant, fractionsArr As Variant
    Dim rempJourArr As Variant, rempNuitArr As Variant
    Dim dateArr As Variant, ferieArr As Variant
    Dim newlyPlaced_presence7_8h As Long
    
    ' === CHARGER CONFIG (utilise Module_Planning_Core) ===
    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    Dim configGlobal As Object
    Set configGlobal = Module_Planning_Core.ChargerConfig(wsConfig)
    
    Dim annee As Long: annee = Module_Planning_Core.CfgLongFromDict(configGlobal, "CFG_Year", Year(Date))

    On Error Resume Next
    nbJours = ws.Cells(4, ws.Columns.count).End(xlToLeft).Column - (colDeb - 1)
    If Err.Number <> 0 Then nbJours = 0
    On Error GoTo 0

    Dim daysInMonth As Long
    On Error Resume Next
    daysInMonth = Day(DateSerial(annee, Module_Planning_Core.MoisNumero(ws.Name) + 1, 0))
    If Err.Number <> 0 Then daysInMonth = 31
    On Error GoTo 0
    
    If nbJours > daysInMonth Then nbJours = daysInMonth
    If nbJours <= 0 Then Exit Sub

    ' === LECTURE TABLEAUX ===
    planningArr = ws.Range(ws.Cells(LIGNE_DEBUT_PLANNING_PERSONNEL, colDeb), ws.Cells(LIGNE_FIN_PLANNING_PERSONNEL, colDeb + nbJours - 1)).Value2
    If LdebFractions > 0 And LfinFractions >= LdebFractions Then
        fractionsArr = ws.Range(ws.Cells(LdebFractions, colDeb), ws.Cells(LfinFractions, colDeb + nbJours - 1)).Value2
    End If
    rempJourArr = ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_JOUR, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_JOUR, colDeb + nbJours - 1)).Value2
    rempNuitArr = ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_NUIT, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_NUIT, colDeb + nbJours - 1)).Value2
    dateArr = ws.Range(ws.Cells(4, colDeb), ws.Cells(4, colDeb + nbJours - 1)).Value2
    ferieArr = ws.Range(ws.Cells(5, colDeb), ws.Cells(5, colDeb + nbJours - 1)).Value2

    If Not IsArray(rempJourArr) Then ReDim rempJourArr(1 To NB_REMPLACEMENT_JOUR_LIGNES, 1 To nbJours)
    If Not IsArray(rempNuitArr) Then ReDim rempNuitArr(1 To NB_REMPLACEMENT_NUIT_LIGNES, 1 To nbJours)

    ' === TABLEAUX CIBLES (statiques pour performance) ===
    Static arrTargetMatin As Variant, arrTargetPM As Variant, arrTargetSoir As Variant
    If IsEmpty(arrTargetMatin) Then
        Dim dataMatin As Variant, dataPM As Variant, dataSoir As Variant
        dataMatin = Array(Array(7, 5), Array(7, 5), Array(7, 5), Array(7, 5), Array(7, 5), Array(5, 5), Array(5, 5))
        arrTargetMatin = CreateTargetArray(dataMatin)
        dataPM = Array(Array(4, 2), Array(3, 2), Array(3, 2), Array(4, 2), Array(4, 2), Array(2, 2), Array(2, 2))
        arrTargetPM = CreateTargetArray(dataPM)
        dataSoir = Array(Array(3, 3), Array(3, 3), Array(3, 3), Array(3, 3), Array(3, 3), Array(3, 3), Array(3, 3))
        arrTargetSoir = CreateTargetArray(dataSoir)
    End If

    ' === CHARGER EFFECTIFS INITIAUX DEPUIS LIGNES 60-62 (OPTIMISATION CLÉ) ===
    Dim initialEffectifsMatin As Variant, initialEffectifsPM As Variant, initialEffectifsSoir As Variant
    initialEffectifsMatin = ws.Range(ws.Cells(60, colDeb), ws.Cells(60, colDeb + nbJours - 1)).Value2
    initialEffectifsPM = ws.Range(ws.Cells(61, colDeb), ws.Cells(61, colDeb + nbJours - 1)).Value2
    initialEffectifsSoir = ws.Range(ws.Cells(62, colDeb), ws.Cells(62, colDeb + nbJours - 1)).Value2
    
    ' === CHARGER JOURS FÉRIÉS (utilise Module_Planning_Core) ===
    Dim joursFeries As Object
    Set joursFeries = Module_Planning_Core.BuildFeriesBE(annee)

    ' === BOUCLE SUR CHAQUE JOUR ===
    For col = 1 To nbJours
        newlyPlaced_presence7_8h = 0

        If IsDate(dateArr(1, col)) Then
            dateJour = CDate(dateArr(1, col))
            jourSemaine = Weekday(dateJour, vbMonday)
        Else
            jourSemaine = ((ws.Cells(4, col + colDeb - 1).Column - colDeb) Mod 7) + 1
        End If
        
        ' Utilise Module_Planning_Core pour détection férié
        codeFerie = Module_Planning_Core.EstDansFeries(dateJour, joursFeries)
        
        Dim targetMatin As Long, targetPM As Long, targetSoir As Long
        targetMatin = arrTargetMatin(jourSemaine - 1, IIf(codeFerie, 1, 0))
        targetPM = arrTargetPM(jourSemaine - 1, IIf(codeFerie, 1, 0))
        targetSoir = arrTargetSoir(jourSemaine - 1, IIf(codeFerie, 1, 0))

        ' === LIRE EFFECTIFS ACTUELS (depuis lignes 60-62) ===
        Dim actualMatin As Long, actualPM As Long, actualSoir As Long
        actualMatin = IIf(IsNumeric(initialEffectifsMatin(1, col)), CLng(initialEffectifsMatin(1, col)), 0)
        actualPM = IIf(IsNumeric(initialEffectifsPM(1, col)), CLng(initialEffectifsPM(1, col)), 0)
        actualSoir = IIf(IsNumeric(initialEffectifsSoir(1, col)), CLng(initialEffectifsSoir(1, col)), 0)

        Dim manqueMatin As Long, manquePM As Long, manqueSoir As Long
        Call ActualiserManquesValeurs(manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir, actualMatin, actualPM, actualSoir)

        ' === REMPLIR REMPLACEMENTS JOUR (lignes 40-41) ===
        For l = 1 To NB_REMPLACEMENT_JOUR_LIGNES
            If Trim(CStr(rempJourArr(l, col))) = "" Then
                Dim code As String: code = ""
                Dim codeSoirTrouve As Boolean: codeSoirTrouve = False
                
                ' PRIORITÉ GÉNÉRALE (Matin+PM)
                If manqueMatin > 0 And manquePM > 0 Then
                    code = codesSuggestion(SUGG_7_1530)(0)
                    If Not CodeDejaPresent(planningArr, rempJourArr, col, code, True) Then
                        rempJourArr(l, col) = code
                        Call MettreAJourCompteursMAS(code, actualMatin, actualPM, actualSoir, LIGNE_REMPLACEMENT_DEBUT_JOUR + l - 1, newlyPlaced_presence7_8h)
                        Call ActualiserManquesValeurs(manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir, actualMatin, actualPM, actualSoir)
                        GoTo NextSlotRemplacement
                    End If
                End If
                
                ' === LOGIQUE SOIR SPÉCIFIQUE (VEN/SAM/FÉRIÉ) ===
                Dim estJourSpecialSoir As Boolean
                estJourSpecialSoir = (jourSemaine = 5 Or jourSemaine = 6 Or codeFerie)

                If estJourSpecialSoir And manqueSoir > 0 Then
                    Dim c19CodePourJourSpecial As String: c19CodePourJourSpecial = "C 19"
                    If codeFerie Then c19CodePourJourSpecial = "C 19 di"

                    Dim c19EstPresentCeJour As Boolean: c19EstPresentCeJour = False
                    Dim c19EstRoleAS As Boolean: c19EstRoleAS = False
                    Dim rScan As Long, valScan As String
                    
                    ' Vérifier C19 dans planning
                    For rScan = LBound(planningArr, 1) To UBound(planningArr, 1)
                        valScan = Trim(CStr(planningArr(rScan, col)))
                        If StrComp(valScan, c19CodePourJourSpecial, vbTextCompare) = 0 Then
                            c19EstPresentCeJour = True
                            If (LIGNE_DEBUT_PLANNING_PERSONNEL + rScan - 1) = LIGNE_AIDE_SOIGNANT_C19_PLANNING Then c19EstRoleAS = True
                            Exit For
                        End If
                    Next rScan

                    ' Vérifier C19 dans remplacements déjà faits
                    If Not c19EstPresentCeJour Then
                        For rScan = 1 To l - 1
                            valScan = Trim(CStr(rempJourArr(rScan, col)))
                            If StrComp(valScan, c19CodePourJourSpecial, vbTextCompare) = 0 Then
                                c19EstPresentCeJour = True
                                c19EstRoleAS = True
                                Exit For
                            End If
                        Next rScan
                    End If
                    
                    Dim nbC20Actuels As Long: nbC20Actuels = 0
                    Dim nbC20EActuels As Long: nbC20EActuels = 0
                    
                    For rScan = LBound(planningArr, 1) To UBound(planningArr, 1)
                        valScan = Trim(CStr(planningArr(rScan, col)))
                        If StrComp(valScan, "C 20", vbTextCompare) = 0 Then nbC20Actuels = nbC20Actuels + 1
                        If StrComp(valScan, "C 20 E", vbTextCompare) = 0 Then nbC20EActuels = nbC20EActuels + 1
                    Next rScan
                    
                    For rScan = 1 To l - 1
                        valScan = Trim(CStr(rempJourArr(rScan, col)))
                        If StrComp(valScan, "C 20", vbTextCompare) = 0 Then nbC20Actuels = nbC20Actuels + 1
                        If StrComp(valScan, "C 20 E", vbTextCompare) = 0 Then nbC20EActuels = nbC20EActuels + 1
                    Next rScan

                    code = ""

                    ' A. Placer C19 si manquant
                    If Not c19EstPresentCeJour And manqueSoir > 0 Then
                        If Not CodeDejaPresent(planningArr, rempJourArr, col, c19CodePourJourSpecial, True) Then
                            rempJourArr(l, col) = c19CodePourJourSpecial
                            Call MettreAJourCompteursMAS(c19CodePourJourSpecial, actualMatin, actualPM, actualSoir, LIGNE_REMPLACEMENT_DEBUT_JOUR + l - 1, newlyPlaced_presence7_8h)
                            Call ActualiserManquesValeurs(manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir, actualMatin, actualPM, actualSoir)
                            c19EstPresentCeJour = True
                            c19EstRoleAS = True
                            codeSoirTrouve = True
                            GoTo NextSlotRemplacement
                        End If
                    End If

                    ' B. Placer C20/C20E selon rôle
                    If c19EstPresentCeJour And manqueSoir > 0 Then
                        If c19EstRoleAS Then
                            If nbC20Actuels = 0 And Not CodeDejaPresent(planningArr, rempJourArr, col, "C 20", True) Then
                                code = "C 20"
                            ElseIf nbC20EActuels = 0 And Not CodeDejaPresent(planningArr, rempJourArr, col, "C 20 E", True) Then
                                code = "C 20 E"
                            End If
                        Else
                            If nbC20Actuels < 2 And Not CodeDejaPresent(planningArr, rempJourArr, col, "C 20", True) Then
                                code = "C 20"
                            End If
                        End If

                        If code <> "" Then
                            rempJourArr(l, col) = code
                            Call MettreAJourCompteursMAS(code, actualMatin, actualPM, actualSoir, LIGNE_REMPLACEMENT_DEBUT_JOUR + l - 1, newlyPlaced_presence7_8h)
                            Call ActualiserManquesValeurs(manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir, actualMatin, actualPM, actualSoir)
                            codeSoirTrouve = True
                            GoTo NextSlotRemplacement
                        End If
                    End If
                    If codeSoirTrouve Then GoTo NextSlotRemplacement

                ' === AUTRES JOURS (manqueSoir) ===
                ElseIf manqueSoir > 0 Then
                    Dim codeSoirGen As String: codeSoirGen = ""
                    Dim c19Gen As String: c19Gen = "C 19"
                    If jourSemaine = 7 Then c19Gen = "C 19 di"

                    If Not CodeDejaPresent(planningArr, rempJourArr, col, c19Gen, True) Then
                        codeSoirGen = c19Gen
                    ElseIf Not CodeDejaPresent(planningArr, rempJourArr, col, "C 20", True) Then
                        codeSoirGen = "C 20"
                    ElseIf Not CodeDejaPresent(planningArr, rempJourArr, col, "C 20 E", True) Then
                        codeSoirGen = "C 20 E"
                    End If
                    
                    If codeSoirGen <> "" Then
                        rempJourArr(l, col) = codeSoirGen
                        Call MettreAJourCompteursMAS(codeSoirGen, actualMatin, actualPM, actualSoir, LIGNE_REMPLACEMENT_DEBUT_JOUR + l - 1, newlyPlaced_presence7_8h)
                        Call ActualiserManquesValeurs(manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir, actualMatin, actualPM, actualSoir)
                        GoTo NextSlotRemplacement
                    End If
                End If

                ' === MANQUE MATIN SEUL ===
                If manqueMatin > 0 Then
                    If Not CodeDejaPresent(planningArr, rempJourArr, col, codesSuggestion(SUGG_7_13)(0), True) Then
                        rempJourArr(l, col) = codesSuggestion(SUGG_7_13)(0)
                        Call MettreAJourCompteursMAS(codesSuggestion(SUGG_7_13)(0), actualMatin, actualPM, actualSoir, LIGNE_REMPLACEMENT_DEBUT_JOUR + l - 1, newlyPlaced_presence7_8h)
                        Call ActualiserManquesValeurs(manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir, actualMatin, actualPM, actualSoir)
                        GoTo NextSlotRemplacement
                    End If
                End If
            End If
NextSlotRemplacement:
        Next l

        ' === REMPLIR REMPLACEMENTS NUIT (lignes 46-47) ===
        Dim nuitCodesProposes As Variant
        If codeFerie Or jourSemaine = 5 Or jourSemaine = 6 Then
            nuitCodesProposes = Array(codesSuggestion(SUGG_NUIT1)(0), codesSuggestion(SUGG_NUIT2)(0))
        Else
            nuitCodesProposes = Array(codesSuggestion(SUGG_NUIT2)(0), codesSuggestion(SUGG_NUIT2)(0))
        End If
        
        Dim iNuitSlot As Long
        For iNuitSlot = 1 To NB_REMPLACEMENT_NUIT_LIGNES
            If UBound(rempNuitArr, 1) >= iNuitSlot And col <= UBound(rempNuitArr, 2) Then
                If Trim(CStr(rempNuitArr(iNuitSlot, col))) = "" Then
                    Dim codeNuitAPlacer As String: codeNuitAPlacer = ""
                    Dim codesPourChoixNuit As Variant
                    
                    If iNuitSlot = 1 Then
                        codesPourChoixNuit = nuitCodesProposes
                        If UBound(nuitCodesProposes) = LBound(nuitCodesProposes) Then
                            codesPourChoixNuit = Array(nuitCodesProposes(LBound(nuitCodesProposes)))
                        End If
                        codeNuitAPlacer = ChoisirCodePertinent(codesPourChoixNuit, planningArr, rempNuitArr, col)
                    Else
                        If CStr(nuitCodesProposes(LBound(nuitCodesProposes))) <> CStr(nuitCodesProposes(UBound(nuitCodesProposes))) Then
                            If Trim(CStr(rempNuitArr(1, col))) = CStr(nuitCodesProposes(LBound(nuitCodesProposes))) Then
                                codeNuitAPlacer = CStr(nuitCodesProposes(UBound(nuitCodesProposes)))
                            ElseIf Trim(CStr(rempNuitArr(1, col))) = CStr(nuitCodesProposes(UBound(nuitCodesProposes))) Then
                                codeNuitAPlacer = CStr(nuitCodesProposes(LBound(nuitCodesProposes)))
                            Else
                                If Not CodeDejaPresent(planningArr, rempNuitArr, col, CStr(nuitCodesProposes(LBound(nuitCodesProposes))), True) Then
                                    codeNuitAPlacer = CStr(nuitCodesProposes(LBound(nuitCodesProposes)))
                                ElseIf Not CodeDejaPresent(planningArr, rempNuitArr, col, CStr(nuitCodesProposes(UBound(nuitCodesProposes))), True) Then
                                    codeNuitAPlacer = CStr(nuitCodesProposes(UBound(nuitCodesProposes)))
                                End If
                            End If
                        Else
                            codeNuitAPlacer = CStr(nuitCodesProposes(LBound(nuitCodesProposes)))
                        End If
                    End If
                    
                    If codeNuitAPlacer <> "" And Not CodeDejaPresent(planningArr, rempNuitArr, col, codeNuitAPlacer, True) Then
                        rempNuitArr(iNuitSlot, col) = codeNuitAPlacer
                    End If
                End If
            End If
        Next iNuitSlot
    Next col

    ' === ÉCRIRE LES TABLEAUX MODIFIÉS ===
    ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_JOUR, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_JOUR, colDeb + nbJours - 1)).Value2 = rempJourArr
    ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_NUIT, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_NUIT, colDeb + nbJours - 1)).Value2 = rempNuitArr
End Sub

' =========================================================================
' MACRO PRINCIPALE
' =========================================================================

Sub AnalyseEtRemplacementPlanningUltraOptimise()
    Dim ws As Worksheet
    Dim LdebFractions As Long, LfinFractions As Long
    Dim colDeb As Long
    Dim choixUtilisateur As VbMsgBoxResult
    Dim codesSuggestion As Variant
    Dim groupesExclusifs As Variant

    colDeb = 3 ' Colonne C (ajuster si besoin : 2=B, 3=C)
    LdebFractions = 0
    LfinFractions = 0

    codesSuggestion = Array( _
        Array("6:45 15:15"), Array("7 15:30"), Array("7 11:30"), Array("7 13"), Array("8 16:30"), _
        Array("C 15", "C 15 di"), Array("C 20"), Array("C 20 E"), Array("C 19"), _
        Array("12:30 16:30"), Array("19:45 6:45"), Array("20 7"))

    groupesExclusifs = Array( _
        Array(codesSuggestion(SUGG_645)(0)), _
        Array(codesSuggestion(SUGG_C15_GRP)(0), codesSuggestion(SUGG_C15_GRP)(1), "C 15di"), _
        Array(codesSuggestion(SUGG_C20_CODE)(0), codesSuggestion(SUGG_C20E_CODE)(0)), _
        Array(codesSuggestion(SUGG_C19_CODE)(0), "C 19 di"))

    choixUtilisateur = MsgBox("Voulez-vous analyser uniquement l'onglet actif (" & ActiveSheet.Name & ") ?" & vbCrLf & _
                              vbCrLf & "Cliquez sur 'Oui' pour l'onglet actif." & vbCrLf & _
                              "Cliquez sur 'Non' pour analyser tous les onglets de mois." & vbCrLf & _
                              "Cliquez sur 'Annuler' pour vider les lignes de remplacement de l'onglet actif.", _
                              vbYesNoCancel + vbQuestion, "Choix de l'analyse")
    
    If colDeb <= 0 Then
        MsgBox "La colonne de début (colDeb = " & colDeb & ") n'est pas valide. Opération annulée.", vbCritical
        Exit Sub
    End If

    If choixUtilisateur = vbCancel Then
        If MsgBox("Voulez-vous VRAIMENT effacer les lignes de remplacement (Lignes " & LIGNE_REMPLACEMENT_DEBUT_JOUR & "-" & LIGNE_REMPLACEMENT_FIN_JOUR & " et " & LIGNE_REMPLACEMENT_DEBUT_NUIT & "-" & LIGNE_REMPLACEMENT_FIN_NUIT & ") de l'onglet '" & ActiveSheet.Name & "'?", _
                  vbYesNo + vbExclamation, "Confirmation Effacement") = vbYes Then
            Set ws = ActiveSheet
            If EstUnOngletDeMois(ws.Name) Then
                Application.ScreenUpdating = False
                Dim lastColData As Long
                On Error Resume Next
                lastColData = ws.Cells(4, ws.Columns.count).End(xlToLeft).Column
                If Err.Number <> 0 Or lastColData < colDeb Then lastColData = colDeb - 1
                On Error GoTo 0

                If lastColData >= colDeb Then
                    ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_JOUR, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_JOUR, lastColData)).ClearContents
                    ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_NUIT, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_NUIT, lastColData)).ClearContents
                End If
                Application.ScreenUpdating = True
                MsgBox "Lignes de remplacement effacées pour l'onglet '" & ws.Name & "'.", vbInformation
            Else
                MsgBox "L'onglet actif (" & ws.Name & ") n'est pas un onglet de mois valide. Effacement non effectué.", vbExclamation
            End If
        Else
            MsgBox "Opération annulée par l'utilisateur.", vbInformation
        End If
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    If choixUtilisateur = vbYes Then
        Set ws = ActiveSheet
        If EstUnOngletDeMois(ws.Name) Then
            Debug.Print "Traitement onglet actif: " & ws.Name
            Call TraiterUneFeuilleDeMois(ws, LdebFractions, LfinFractions, colDeb, groupesExclusifs, codesSuggestion)
            MsgBox "Analyse et remplacements pour l'onglet '" & ws.Name & "' terminés !", vbInformation
        Else
            MsgBox "L'onglet actif (" & ws.Name & ") n'est pas un onglet de mois valide. Opération non effectuée.", vbExclamation
        End If
    Else
        For Each ws In ThisWorkbook.Worksheets
            If EstUnOngletDeMois(ws.Name) Then
                Debug.Print "Traitement onglet: " & ws.Name
                Call TraiterUneFeuilleDeMois(ws, LdebFractions, LfinFractions, colDeb, groupesExclusifs, codesSuggestion)
            End If
        Next ws
        MsgBox "Analyse et remplacements pour tous les mois terminés !", vbInformation
    End If

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Debug.Print "Fin du script."
    Exit Sub

ErrorHandler:
    MsgBox "Erreur d'exécution N° " & Err.Number & ":" & vbCrLf & Err.description & vbCrLf & "Source: " & Err.Source, vbCritical, "Erreur VBA"
    Resume CleanExit
End Sub


