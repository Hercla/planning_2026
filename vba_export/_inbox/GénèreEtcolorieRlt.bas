' ExportedAt: 2026-01-12 15:37:08 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "GénèreEtcolorieRlt"
'==================================================================
'              Génération et Coloration du Roulement Optimisé
'  Utilise "Feuil_Config" pour couleurs des codes.
'
'  Fonctionnalités :
'   - Copie le roulement 56 jrs vers Janv–Déc
'   - Respecte Jour / Nuit (séparateur "Nuit" détecté avec tolérance)
'   - Garde les couleurs du roulement
'   - Si pas de couleur ? va chercher dans Feuil_Config (CO/CP)
'   - Collage forcé en TEXTE (pas de conversion heure)
'   - Police :
'       * EXACTEMENT "8:30 12:45 16:30 20:15" = Arial 8 non gras
'       * Tout le reste (WE, C 15, 7 15:30, DP 7:13, …) = Arial 12 non gras
'   - Hauteur de ligne FIXE à 45 points (60 pixels) pour toutes les lignes
'==================================================================

' --- CONSTANTES ---
Private Const ROULEMENT_SHEET_NAME As String = "Roulements"
Private Const ACCEUIL_SHEET_NAME   As String = "Feuil_Config"

Private Const PERSONNEL_START_ROW  As Long = 6
Private Const PERSONNEL_NAME_COL   As String = "B"
Private Const ROULEMENT_START_COL  As Long = 4    ' colonne D

Private Const NUIT_SEPARATOR       As String = "Nuit" ' (conservé pour cohérence)

' colonnes dans Feuil_Config
Private Const CFG_COL_CODE         As String = "CO"
Private Const CFG_COL_COLOR        As String = "CP"
Private Const CFG_FIRST_ROW        As Long = 2

' zones dans les feuilles mois
Private Const TARGET_NAME_COL          As String = "A"
Private Const TARGET_START_ROW_JOUR    As Long = 6
Private Const TARGET_END_ROW_JOUR      As Long = 28
Private Const TARGET_START_ROW_NUIT    As Long = 31
Private Const TARGET_END_ROW_NUIT      As Long = 38

' 1er jour = colonne C
Private Const TARGET_START_COL     As Long = 3
Private Const TARGET_END_COL       As Long = 33

' --- VARS GLOBALES ---
Private dictColorCache As Object

' ==========================================================
' MACRO PRINCIPALE
' ==========================================================
Public Sub GenererRoulementOptimise()

    Dim wsRoulement As Worksheet, wsConfig As Worksheet
    Dim dateDebut As Date, finAnnee As Date
    Dim lastPersRow As Long, nightRow As Long
    Dim dictCouleurs As Object, dictLignesCibles As Object
    Dim choixTous As VbMsgBoxResult, filtreNom As String
    Dim moisSheets(1 To 12) As Worksheet
    Dim joursCol(1 To 56) As Long
    Dim jourDepart As Long, jourInput As String
    Dim colDepart As Long, joursRestantsDansRoulement As Long
    Dim dateCourante As Date, joursRestants As Long, blocJours As Long
    Dim roulementData As Variant, roulementInteriorColors As Variant, roulementFontColors As Variant
    Dim ligneEmploye As Long, offsetCol As Long
    Dim targetRow As Long, targetCol As Long
    Dim targetSheet As Worksheet
    Dim moisIndex As Long
    Dim equipe As String
    Dim currentDateInBlock As Date
    Dim nomEmploye As String

On Error GoTo ErrorHandler

    ' 1. préparer environnement + charger couleurs
    If Not SetupEnvironment(wsRoulement, wsConfig, dictCouleurs) Then GoTo CleanUp

    ' 2. paramètres user
    If Not GetUserParameters(dateDebut, finAnnee, choixTous, filtreNom) Then GoTo CleanUp

    ' 3. structure de l'onglet roulement (Nuit tolérant)
    If Not FindLayoutInfo(wsRoulement, nightRow, lastPersRow) Then GoTo CleanUp

    ' 4. charger les 12 mois
    LoadMonthlySheets moisSheets

    ' 5. mapping nom+équipe+mois -> ligne dans le mois (ROBUSTE)
    Set dictLignesCibles = CreateTargetLineDictionary( _
        moisSheets, wsRoulement, PERSONNEL_START_ROW, lastPersRow, nightRow _
    )
    If dictLignesCibles Is Nothing Then
        MsgBox "Erreur lors de la création du mapping des employés.", vbCritical
        GoTo CleanUp
    End If

    ' 6. pré-calcul colonnes 1..56 dans Roulements
    PrecomputeJoursColonnes wsRoulement, joursCol

    ' 7. demander le jour de départ
    jourInput = InputBox("Entrez le numéro de jour de départ (1 à 56) :", _
                         "Jour de départ du roulement", 1)
    If Not IsNumeric(jourInput) Or val(jourInput) < 1 Or val(jourInput) > 56 Then
        MsgBox "Numéro de jour invalide (1..56).", vbCritical
        GoTo CleanUp
    End If
    jourDepart = CLng(jourInput)

    ' 8. déroulage
    dateCourante = dateDebut

    Do While dateCourante <= finAnnee

        colDepart = joursCol(jourDepart)
        joursRestantsDansRoulement = 56 - (jourDepart - 1)
        joursRestants = finAnnee - dateCourante + 1
        blocJours = Application.Min(joursRestants, joursRestantsDansRoulement)

        ' récupérer bloc source
        With wsRoulement.Range( _
            wsRoulement.Cells(PERSONNEL_START_ROW, colDepart), _
            wsRoulement.Cells(lastPersRow, colDepart + blocJours - 1) _
        )
            roulementData = .value
            GetRangeColors .Cells, roulementInteriorColors, roulementFontColors
        End With

        ' parcourir chaque ligne (= chaque employé)
        For ligneEmploye = 1 To UBound(roulementData, 1)

            Dim srcRow As Long
            srcRow = ligneEmploye + PERSONNEL_START_ROW - 1

            ' sauter la ligne "Nuit"
            If srcRow = nightRow Then GoTo NextEmployee

            nomEmploye = NormalizeName(wsRoulement.Cells(srcRow, PERSONNEL_NAME_COL).value)
            If nomEmploye = "" Then GoTo NextEmployee

            ' filtre éventuel
            If choixTous = vbNo Then
                If InStr(1, nomEmploye, NormalizeName(filtreNom), vbTextCompare) = 0 Then GoTo NextEmployee
            End If

            ' équipe selon position
            equipe = IIf(srcRow < nightRow, "Jour", "Nuit")
            currentDateInBlock = dateCourante

            ' coller jour par jour pour cet employé
            For offsetCol = 1 To blocJours

                If currentDateInBlock > finAnnee Then Exit For

                Dim horaire As String
                Dim srcBackColor As Long, srcFontColor As Long

                horaire = CStr(roulementData(ligneEmploye, offsetCol))
                horaire = NormalizeTextCell(horaire) ' normalise les espaces/retours
                srcBackColor = roulementInteriorColors(ligneEmploye, offsetCol)
                srcFontColor = roulementFontColors(ligneEmploye, offsetCol)

                If horaire <> "" Then
                    moisIndex = Month(currentDateInBlock)
                    Set targetSheet = moisSheets(moisIndex)

                    If Not targetSheet Is Nothing Then
                        Dim keyLookup As String
                        keyLookup = nomEmploye & "|" & equipe & "|" & moisIndex

                        If dictLignesCibles.Exists(keyLookup) Then
                            targetRow = dictLignesCibles(keyLookup)

                            If targetRow > 0 Then
                                targetCol = Day(currentDateInBlock) + TARGET_START_COL - 1

                                With targetSheet.Cells(targetRow, targetCol)
                                    ' --- FORCER TEXTE avant d'écrire ---
                                    .NumberFormat = "@"
                                    .Formula = "'" & horaire   ' colle en texte pur

                                    ' --- COULEUR DE FOND ---
                                    If srcBackColor <> 0 Then
                                        .Interior.Color = srcBackColor                   ' priorité : couleur du roulement
                                    Else
                                        ' pas de couleur dans le roulement ? on tente Feuil_Config
                                        If Not dictCouleurs Is Nothing Then
                                            If dictCouleurs.Exists(horaire) Then
                                                If dictCouleurs(horaire) <> 0 Then
                                                    .Interior.Color = dictCouleurs(horaire)
                                                Else
                                                    .Interior.ColorIndex = xlColorIndexNone
                                                End If
                                            Else
                                                .Interior.ColorIndex = xlColorIndexNone
                                            End If
                                        Else
                                            .Interior.ColorIndex = xlColorIndexNone
                                        End If
                                    End If

                                    ' --- COULEUR DE POLICE ---
                                    ' Par défaut : noir. Puis on écrase si une couleur source existe.
                                    .Font.Color = vbBlack
                                    If srcFontColor <> 0 Then
                                        .Font.Color = srcFontColor
                                    End If

                                    ' --- mise en forme commune ---
                                    .WrapText = True
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlCenter
                                    .Font.Bold = False
                                    .Font.Name = "Arial"

                                    ' --- taille ciblée ---
                                    If IsSpecialFourTimes(horaire) Then
                                        .Font.Size = 8     ' uniquement "8:30 12:45 16:30 20:15"
                                    Else
                                        .Font.Size = 12    ' tout le reste
                                    End If
                                End With

                                ' CORRECTION : hauteur FIXE de 45 points (60 pixels) au lieu de AutoFit
                                targetSheet.Rows(targetRow).RowHeight = 45
                            End If
                        End If
                    End If
                End If

                currentDateInBlock = currentDateInBlock + 1
            Next offsetCol

NextEmployee:
        Next ligneEmploye

        ' avancer dans le calendrier réel
        dateCourante = dateCourante + blocJours

        ' avancer dans le cycle 56j
        If jourDepart + blocJours - 1 >= 56 Then
            jourDepart = 1
        Else
            jourDepart = jourDepart + blocJours
        End If

        ' nettoyer les arrays
        Erase roulementData
        Erase roulementInteriorColors
        Erase roulementFontColors

    Loop

    MsgBox "Roulement collé (couleurs + police dynamique, texte forcé).", vbInformation

CleanUp:
    RestoreEnvironment
    Exit Sub

ErrorHandler:
    MsgBox "Erreur dans GenererRoulementOptimise : " & vbCrLf & Err.Description, vbCritical
    Resume CleanUp
End Sub

' ==========================================================
' Pré-calcul des colonnes jour 1..56
' ==========================================================
Private Sub PrecomputeJoursColonnes(ws As Worksheet, ByRef joursCol() As Long)
    Dim i As Long
    For i = 1 To 56
        joursCol(i) = ws.Cells(3, ROULEMENT_START_COL + i - 1).Column
    Next i
End Sub

' ==========================================================
' Préparation appli + chargement dico
' ==========================================================
Private Function SetupEnvironment(ByRef wsRoul As Worksheet, _
                                  ByRef wsAcc As Worksheet, _
                                  ByRef dictCol As Object) As Boolean
On Error Resume Next
    Set wsRoul = ThisWorkbook.Sheets(ROULEMENT_SHEET_NAME)
    Set wsAcc = ThisWorkbook.Sheets(ACCEUIL_SHEET_NAME)
On Error GoTo 0

    If wsRoul Is Nothing Or wsAcc Is Nothing Then
        MsgBox "Feuilles '" & ROULEMENT_SHEET_NAME & "' ou '" & ACCEUIL_SHEET_NAME & "' introuvables.", vbCritical
        SetupEnvironment = False
        Exit Function
    End If

    ' charger les couleurs depuis Feuil_Config
    Set dictCol = LoadColorsFromAcceuil(wsAcc)
    Set dictColorCache = dictCol

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    SetupEnvironment = True
End Function

Private Sub RestoreEnvironment()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
    End With
End Sub

' ==========================================================
' LECTURE DES COULEURS DANS FEUIL_CONFIG
' ==========================================================
Private Function LoadColorsFromAcceuil(wsAcc As Worksheet) As Object
    Dim d As Object
    Dim r As Long
    Dim codeVal As String
    Dim colorVal As Long
    Dim lastRow As Long
    Dim altColor As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    lastRow = wsAcc.Cells(wsAcc.Rows.Count, CFG_COL_CODE).End(xlUp).row
    If lastRow < CFG_FIRST_ROW Then lastRow = CFG_FIRST_ROW

    For r = CFG_FIRST_ROW To lastRow
        codeVal = NormalizeTextCell(wsAcc.Cells(r, CFG_COL_CODE).value)
        If codeVal <> "" Then
            colorVal = wsAcc.Cells(r, CFG_COL_COLOR).Interior.Color
            If colorVal = 0 Or colorVal = vbWhite Then
                altColor = wsAcc.Cells(r, CFG_COL_CODE).Interior.Color
                If altColor <> 0 Then colorVal = altColor
            End If
            d(codeVal) = colorVal
        End If
    Next r

    Set LoadColorsFromAcceuil = d
End Function

' ==========================================================
' Paramètres utilisateur
' ==========================================================
Private Function GetUserParameters(ByRef startDate As Date, ByRef endDate As Date, _
                                   ByRef allEmpl As VbMsgBoxResult, ByRef nameFilter As String) As Boolean
    Dim dateInput As String

    dateInput = InputBox("Entrez le lundi de départ (format JJ/MM/AAAA) :", "Date Début Roulement")
    If Not IsDate(dateInput) Then
        MsgBox "Date invalide.", vbCritical
        GetUserParameters = False
        Exit Function
    End If

    startDate = CDate(dateInput)
    If Weekday(startDate, vbMonday) <> 1 Then
        MsgBox "La date de départ doit être un lundi.", vbCritical
        GetUserParameters = False
        Exit Function
    End If

    endDate = DateSerial(Year(startDate), 12, 31)

    allEmpl = MsgBox("Appliquer le roulement à tout le personnel ?", _
                     vbYesNoCancel + vbQuestion, "Choix Personnel")
    If allEmpl = vbCancel Then
        GetUserParameters = False
        Exit Function
    End If

    nameFilter = ""
    If allEmpl = vbNo Then
        nameFilter = InputBox("Entrez une partie du nom/prénom du membre du personnel :", "Filtrer Personnel")
        If Trim$(nameFilter) = "" Then
            MsgBox "Aucun nom entré pour le filtre.", vbExclamation
            GetUserParameters = False
            Exit Function
        End If
    End If

    GetUserParameters = True
End Function

' ==========================================================
' Cherche la ligne "Nuit" (tolérant) + dernière ligne employé
' ==========================================================
Private Function FindLayoutInfo(ByVal ws As Worksheet, _
                                ByRef nightRow As Long, _
                                ByRef lastRow As Long) As Boolean
    Dim r As Long, c As Long, usedCols As Long
    Dim txt As String, guess As Variant

    FindLayoutInfo = False
    nightRow = 0
    lastRow = ws.Cells(ws.Rows.Count, PERSONNEL_NAME_COL).End(xlUp).row
    If lastRow < PERSONNEL_START_ROW Then
        MsgBox "Aucun personnel trouvé sur '" & ws.Name & "'.", vbCritical
        Exit Function
    End If

    usedCols = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If usedCols < 3 Then usedCols = 3

    ' 1) recherche stricte, colonne des noms (B)
    For r = PERSONNEL_START_ROW To lastRow
        txt = NormalizeTextCell(ws.Cells(r, PERSONNEL_NAME_COL).value)
        If IsWordNuit(txt) Then nightRow = r: Exit For
    Next r

    ' 2) élargir à la colonne A si pas trouvé
    If nightRow = 0 Then
        For r = PERSONNEL_START_ROW To lastRow
            txt = NormalizeTextCell(ws.Cells(r, "A").value)
            If IsWordNuit(txt) Then nightRow = r: Exit For
        Next r
    End If

    ' 3) élargir aux premières colonnes A:E
    If nightRow = 0 Then
        For r = PERSONNEL_START_ROW To lastRow
            For c = 1 To Application.Min(usedCols, 5)
                txt = NormalizeTextCell(ws.Cells(r, c).value)
                If IsWordNuit(txt) Then nightRow = r: Exit For
            Next c
            If nightRow <> 0 Then Exit For
        Next r
    End If

    ' 4) si toujours pas trouvé, demander à l'utilisateur
    If nightRow = 0 Then
        guess = InputBox("Ligne 'Nuit' introuvable. Entrez le numéro de ligne du séparateur 'Nuit' (ou Annuler).", _
                         "Indiquer la ligne 'Nuit'")
        If IsNumeric(guess) Then
            nightRow = CLng(guess)
            If nightRow < PERSONNEL_START_ROW Or nightRow > lastRow Then
                MsgBox "Numéro de ligne invalide pour 'Nuit'.", vbCritical
                Exit Function
            End If
        Else
            MsgBox "Le séparateur 'Nuit' est introuvable.", vbExclamation
            Exit Function
        End If
    End If

    FindLayoutInfo = True
End Function

Private Function IsWordNuit(ByVal s As String) As Boolean
    Dim t As String
    t = LCase$(s)
    t = Replace(t, "_", " ")
    t = Replace(t, "-", " ")
    t = Replace(t, "—", " ")
    t = WorksheetFunction.Trim(t)
    IsWordNuit = (t = "nuit") Or (InStr(1, " " & t & " ", " nuit ") > 0)
End Function

' ==========================================================
' Charger les feuilles de mois
' ==========================================================
Private Sub LoadMonthlySheets(ByRef sheetArray() As Worksheet)
    Dim monthNames As Variant
    Dim i As Long
    monthNames = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                       "Juillet", "Aout", "Sept", "Oct", "Nov", "Dec")
    For i = 0 To 11
        On Error Resume Next
        Set sheetArray(i + 1) = ThisWorkbook.Sheets(monthNames(i))
        On Error GoTo 0
    Next i
End Sub

' ==========================================================
' Dictionnaire employé/mois -> ligne dans feuille mois (ROBUSTE)
' ==========================================================
Private Function CreateTargetLineDictionary(ByRef monthlySheets() As Worksheet, _
                                           wsRoul As Worksheet, _
                                           startR As Long, endR As Long, nightR As Long) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim monthIndex As Long, wsM As Worksheet
    Dim idxJour As Object, idxNuit As Object
    Dim r As Long, empKey As String, equipe As String, key As String, tgt As Long

    For monthIndex = 1 To 12
        Set wsM = monthlySheets(monthIndex)
        If Not wsM Is Nothing Then
            Set idxJour = BuildNameIndex(wsM, TARGET_START_ROW_JOUR, TARGET_END_ROW_JOUR)
            Set idxNuit = BuildNameIndex(wsM, TARGET_START_ROW_NUIT, TARGET_END_ROW_NUIT)
        Else
            Set idxJour = CreateObject("Scripting.Dictionary")
            Set idxNuit = CreateObject("Scripting.Dictionary")
        End If

        For r = startR To endR
            If r = nightR Then GoTo NextRow

            empKey = NormalizeName(wsRoul.Cells(r, PERSONNEL_NAME_COL).value)
            If empKey = "" Then GoTo NextRow

            equipe = IIf(r < nightR, "Jour", "Nuit")

            If equipe = "Jour" Then
                tgt = ResolveTargetRow(empKey, idxJour)
                If tgt = 0 Then tgt = ResolveTargetRow(empKey, idxNuit) ' filet de sécurité
            Else
                tgt = ResolveTargetRow(empKey, idxNuit)
                If tgt = 0 Then tgt = ResolveTargetRow(empKey, idxJour)
            End If

            key = empKey & "|" & equipe & "|" & monthIndex
            dict(key) = tgt

NextRow:
        Next r
    Next monthIndex

    Set CreateTargetLineDictionary = dict
End Function

Private Function BuildNameIndex(ws As Worksheet, rowStart As Long, rowEnd As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim r As Long, v As String
    For r = rowStart To rowEnd
        v = NormalizeName(ws.Cells(r, TARGET_NAME_COL).value)
        If v <> "" Then d(v) = r
    Next r
    Set BuildNameIndex = d
End Function

Private Function ResolveTargetRow(empKey As String, idx As Object) As Long
    Dim k As Variant
    If idx.Exists(empKey) Then ResolveTargetRow = idx(empKey): Exit Function
    ' match partiel si exact introuvable
    For Each k In idx.Keys
        If InStr(k, empKey) > 0 Or InStr(empKey, k) > 0 Then
            ResolveTargetRow = idx(k): Exit Function
        End If
    Next k
    ResolveTargetRow = 0
End Function

' ==========================================================
' Lit les couleurs d'un Range dans 2 arrays
' ==========================================================
Private Sub GetRangeColors(ByVal rng As Range, _
                           ByRef interiorColors As Variant, _
                           ByRef fontColors As Variant)
    Dim r As Long, c As Long
    Dim rowsCount As Long, colsCount As Long

    rowsCount = rng.Rows.Count
    colsCount = rng.Columns.Count

    ReDim interiorColors(1 To rowsCount, 1 To colsCount)
    ReDim fontColors(1 To rowsCount, 1 To colsCount)

    For r = 1 To rowsCount
        For c = 1 To colsCount
            interiorColors(r, c) = rng.Cells(r, c).Interior.Color
            fontColors(r, c) = rng.Cells(r, c).Font.Color
        Next c
    Next r
End Sub

' ==========================================================
' Utils de normalisation + détection bloc heures
' ==========================================================
Private Function NormalizeName(ByVal s As String) As String
    s = Replace(s, Chr(160), " ") ' espace insécable ? espace
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = WorksheetFunction.Trim(s) ' compresse les espaces
    NormalizeName = LCase$(s)     ' insensible casse
End Function

Private Function NormalizeTextCell(ByVal s As String) As String
    s = Replace(CStr(s), Chr(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    NormalizeTextCell = WorksheetFunction.Trim(s)
End Function

' EXACTEMENT le bloc 4 horaires ? font 8, sinon 12
Private Function IsSpecialFourTimes(ByVal txt As String) As Boolean
    IsSpecialFourTimes = (NormalizeTextCell(txt) = "8:30 12:45 16:30 20:15")
End Function



