Attribute VB_Name = "Module_Calculer_Totaux"
'Attribute VB_Name = "Module_Calculer_Totaux"
Option Explicit
' =============================================================================
' MODULE : Module_Calculer_Totaux
' VERSION : 5.0 - FIX CODES NUIT + FORMAT "X (Y)"
' DATE : 2026-02-04
' DESCRIPTION : Calcul des totaux de presence et mise a jour du cockpit
' NOUVEAUTE v5: Fix matching codes nuit (apostrophe), lignes 74-76 remplies
' =============================================================================

' --- CONSTANTES GLOBALES ---
Private Const SEUIL_MATIN As Double = 13
Private Const SEUIL_PM As Double = 13
Private Const SEUIL_SOIR_DEMI As Double = 16.5
Private Const SEUIL_SOIR_PLEIN As Double = 17.5
Private Const SEUIL_NUIT_DEBUT As Double = 19.5
Private Const SEUIL_NUIT_FIN As Double = 7.25

Sub Calculer_Totaux_Planning()
    Dim ws As Worksheet
    Dim wsCodesSpec As Worksheet, wsConfigCodes As Worksheet, wsConfig As Worksheet
    Dim startTime As Double: startTime = Timer

    Set ws = ActiveSheet
    
    ' --- VALIDATION ONGLET ---
    Dim nomOnglet As String: nomOnglet = ws.Name
    If Not IsMonthSheet(nomOnglet) Then
        MsgBox "ERREUR : Selectionnez un onglet mois (Janv-Dec)", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsCodesSpec = ThisWorkbook.Sheets("Codes_Speciaux")
    Set wsConfigCodes = ThisWorkbook.Sheets("Config_Codes")
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    If wsCodesSpec Is Nothing And wsConfigCodes Is Nothing Then
        MsgBox "ERREUR : Feuilles config introuvables.", vbCritical: Exit Sub
    End If

    ' === CHARGEMENT CONFIG (1 seule fois) ===
    Dim cfg As Object: Set cfg = ChargerConfigFast(wsConfig)
    
    Dim ligneDebut As Long: ligneDebut = GCL(cfg, "CHK_FirstPersonnelRow", 6)
    Dim ligneFin As Long: ligneFin = GCL(cfg, "ligneFin", 28)
    Dim colDebut As Long: colDebut = GCL(cfg, "PLN_FirstDayCol", 3)
    Dim colFin As Long: colFin = GCL(cfg, "PLN_LastDayCol", 33)
    Dim ligneNumJour As Long: ligneNumJour = GCL(cfg, "PLN_Row_DayNumbers", 4)
    Dim couleurIgnore As Long: couleurIgnore = GCL(cfg, "CHK_IgnoreColor", 15849925)
    Dim couleurINFAdmin As Long: couleurINFAdmin = GCL(cfg, "COULEUR_INF_ADMIN", 65535)
    Dim couleurBleuClair As Long: couleurBleuClair = GCL(cfg, "COULEUR_BLEU_CLAIR", 15128749)
    
    Dim fonctionsACompter As String: fonctionsACompter = "," & UCase(Replace(Replace(GCS(cfg, "CHK_InfFunctions"), " ", ""), ";", ",")) & ","
    If fonctionsACompter = ",," Then fonctionsACompter = ",INF,AS,CEFA,"
    
    ' --- COCKPIT ROWS ---
    Dim rDates As Long: rDates = GCL(cfg, "COCKPIT_ROW_DATES", 59)
    Dim rDatesSrc As Long: rDatesSrc = GCL(cfg, "COCKPIT_ROW_DATES_SRC", 4)
    Dim rLabelStart As Long: rLabelStart = GCL(cfg, "COCKPIT_ROW_LABEL_START", 59)
    Dim rLabelEnd As Long: rLabelEnd = GCL(cfg, "COCKPIT_ROW_LABEL_END", 78)
    Dim rMeteo As Long: rMeteo = GCL(cfg, "COCKPIT_ROW_METEO", 63)
    Dim rMatin As Long: rMatin = GCL(cfg, "COCKPIT_ROW_MATIN", 64)
    Dim r7h8h As Long: r7h8h = GCL(cfg, "COCKPIT_ROW_7H8H", 65)
    Dim rAM As Long: rAM = GCL(cfg, "COCKPIT_ROW_AM", 66)
    Dim rSoir As Long: rSoir = GCL(cfg, "COCKPIT_ROW_SOIR", 67)
    Dim rNuit As Long: rNuit = GCL(cfg, "COCKPIT_ROW_NUIT", 68)
    Dim rP0645 As Long: rP0645 = GCL(cfg, "COCKPIT_ROW_P0645", 67)
    Dim rP7H8H As Long: rP7H8H = GCL(cfg, "COCKPIT_ROW_P7H8H", 68)
    Dim rP81630 As Long: rP81630 = GCL(cfg, "COCKPIT_ROW_P81630", 69)
    Dim rC15 As Long: rC15 = GCL(cfg, "COCKPIT_ROW_C15", 70)
    Dim rC20 As Long: rC20 = GCL(cfg, "COCKPIT_ROW_C20", 71)
    Dim rC20E As Long: rC20E = GCL(cfg, "COCKPIT_ROW_C20E", 72)
    Dim rC19 As Long: rC19 = GCL(cfg, "COCKPIT_ROW_C19", 73)
    Dim rNuitCode1 As Long: rNuitCode1 = GCL(cfg, "COCKPIT_ROW_NUIT_CODE1", 74)
    Dim rNuitCode2 As Long: rNuitCode2 = GCL(cfg, "COCKPIT_ROW_NUIT_CODE2", 75)
    Dim rNuitTotal As Long: rNuitTotal = GCL(cfg, "COCKPIT_ROW_NUIT_TOTAL", 76)
    Dim rAction As Long: rAction = GCL(cfg, "COCKPIT_ROW_ACTION", 77)
    Dim lig1 As Long: lig1 = GCL(cfg, "CALC_ROW_Matin", 60)
    Dim lig2 As Long: lig2 = GCL(cfg, "CALC_ROW_AM", 61)
    Dim lig3 As Long: lig3 = GCL(cfg, "CALC_ROW_Soir", 62)

    ' --- EFFECTIFS ---
    Dim effSem(1 To 4) As Long, effWE(1 To 4) As Long, effFER(1 To 4) As Long
    effSem(1) = GCL(cfg, "EFF_SEM_Matin", 7): effSem(2) = GCL(cfg, "EFF_SEM_PM", 3)
    effSem(3) = GCL(cfg, "EFF_SEM_Soir", 3): effSem(4) = GCL(cfg, "EFF_SEM_Nuit", 2)
    effWE(1) = GCL(cfg, "EFF_WE_Matin", 5): effWE(2) = GCL(cfg, "EFF_WE_PM", 2)
    effWE(3) = GCL(cfg, "EFF_WE_Soir", 3): effWE(4) = GCL(cfg, "EFF_WE_Nuit", 2)
    effFER(1) = GCL(cfg, "EFF_FER_Matin", 5): effFER(2) = GCL(cfg, "EFF_FER_PM", 2)
    effFER(3) = GCL(cfg, "EFF_FER_Soir", 3): effFER(4) = GCL(cfg, "EFF_FER_Nuit", 2)

    ' --- AUTRES CONFIG ---
    Dim cfg78MinTotal As Long: cfg78MinTotal = GCL(cfg, "COCKPIT_7H8H_MIN_TOTAL", 2)
    Dim cfg78MinInf As Long: cfg78MinInf = GCL(cfg, "COCKPIT_7H8H_MIN_INF", 2)
    Dim cfg78MinAS As Long: cfg78MinAS = GCL(cfg, "COCKPIT_7H8H_MIN_AS", 1)
    Dim cfg78AllowAllInf As Long: cfg78AllowAllInf = GCL(cfg, "COCKPIT_7H8H_ALLOW_ALL_INF", 1)
    Dim cfgP0645Req As Long: cfgP0645Req = GCL(cfg, "COCKPIT_P0645_REQUIRED", 1)
    Dim cfgC19Req As Long: cfgC19Req = GCL(cfg, "COCKPIT_C19_REQUIRED", 1)
    Dim cfgC15ReqCount As Long: cfgC15ReqCount = GCL(cfg, "COCKPIT_C15_REQUIRED_COUNT", 1)
    Dim cfgC20ReqSpecial As Long: cfgC20ReqSpecial = GCL(cfg, "COCKPIT_C20_REQUIRED_SPECIAL", 2)
    Dim cfgCoupeMinTotal As Long: cfgCoupeMinTotal = GCL(cfg, "COCKPIT_COUPES_MIN_TOTAL", 3)
    Dim cfgNuitReqTotal As Long: cfgNuitReqTotal = GCL(cfg, "COCKPIT_NUIT_REQUIRED_TOTAL", 2)
    Dim cfgNuitCode1 As String: cfgNuitCode1 = GCS(cfg, "COCKPIT_NUIT_CODE_1"): If cfgNuitCode1 = "" Then cfgNuitCode1 = "19:45 6:45"
    Dim cfgNuitCode2 As String: cfgNuitCode2 = GCS(cfg, "COCKPIT_NUIT_CODE_2"): If cfgNuitCode2 = "" Then cfgNuitCode2 = "20 7"
    Dim cfgSpecialC20 As String: cfgSpecialC20 = UCase(GCS(cfg, "COCKPIT_SPECIAL_C20_DAYS")): If cfgSpecialC20 = "" Then cfgSpecialC20 = "VEN,SAM,FERIE"
    Dim cfgC15Forbidden As String: cfgC15Forbidden = UCase(GCS(cfg, "COCKPIT_C15_FORBIDDEN_DAYS")): If cfgC15Forbidden = "" Then cfgC15Forbidden = "VEN,SAM,FERIE"
    Dim cfgC15Required As String: cfgC15Required = UCase(GCS(cfg, "COCKPIT_C15_REQUIRED_DAYS")): If cfgC15Required = "" Then cfgC15Required = "LUN,MAR,MER,JEU,DIM"
    
    ' --- NORMALISER CODES NUIT CONFIG (pour matching) ---
    Dim cfgNuitCode1Norm As String: cfgNuitCode1Norm = NormaliserCodeNuit(cfgNuitCode1)
    Dim cfgNuitCode2Norm As String: cfgNuitCode2Norm = NormaliserCodeNuit(cfgNuitCode2)
    
    ' --- FONTS ---
    Dim cfgDeltaFontSize As Long: cfgDeltaFontSize = GCL(cfg, "COCKPIT_DELTA_FONT_SIZE", 16)
    Dim cfgDeltaFontName As String: cfgDeltaFontName = GCS(cfg, "COCKPIT_DELTA_FONT_NAME"): If cfgDeltaFontName = "" Then cfgDeltaFontName = "Arial Narrow"
    Dim cfgCheckFontOk As Long: cfgCheckFontOk = GCL(cfg, "COCKPIT_CHECK_FONT_SIZE_OK", 14)
    Dim cfgCheckFontAlert As Long: cfgCheckFontAlert = GCL(cfg, "COCKPIT_CHECK_FONT_SIZE_ALERT", 9)
    Dim cfgCheckFontName As String: cfgCheckFontName = GCS(cfg, "COCKPIT_CHECK_FONT_NAME"): If cfgCheckFontName = "" Then cfgCheckFontName = "Arial Narrow"
    Dim cfgLabelFontName As String: cfgLabelFontName = GCS(cfg, "COCKPIT_LABEL_FONT_NAME"): If cfgLabelFontName = "" Then cfgLabelFontName = "Arial Narrow"
    Dim cfgLabelFontSize As Long: cfgLabelFontSize = GCL(cfg, "COCKPIT_LABEL_FONT_SIZE", 16)
    Dim cfgLabelFontBold As Long: cfgLabelFontBold = GCL(cfg, "COCKPIT_LABEL_FONT_BOLD", 0)
    Dim cfgTotalFontName As String: cfgTotalFontName = GCS(cfg, "COCKPIT_TOTAL_FONT_NAME"): If cfgTotalFontName = "" Then cfgTotalFontName = cfgCheckFontName
    Dim cfgTotalFontSize As Long: cfgTotalFontSize = GCL(cfg, "COCKPIT_TOTAL_FONT_SIZE", cfgCheckFontOk)
    Dim cfgTotalOkColor As Long: cfgTotalOkColor = GCL(cfg, "COCKPIT_TOTAL_OK_COLOR", 13434828)
    Dim cfgTotalBadColor As Long: cfgTotalBadColor = GCL(cfg, "COCKPIT_TOTAL_BAD_COLOR", 6710886)
    Dim cfgTotalOkFontColor As Long: cfgTotalOkFontColor = GCL(cfg, "COCKPIT_TOTAL_OK_FONT_COLOR", 7895160)
    Dim cfgTotalBadFontColor As Long: cfgTotalBadFontColor = GCL(cfg, "COCKPIT_TOTAL_BAD_FONT_COLOR", 16777215)
    Dim cfgTotalLabelFontSize As Long: cfgTotalLabelFontSize = GCL(cfg, "COCKPIT_TOTAL_LABEL_SIZE", cfgLabelFontSize)
    Dim cfgTotalLabelFontColor As Long: cfgTotalLabelFontColor = GCL(cfg, "COCKPIT_TOTAL_LABEL_COLOR", 0)
    Dim cfgTotalRowHeight As Long: cfgTotalRowHeight = GCL(cfg, "COCKPIT_TOTAL_ROW_HEIGHT", 18)
    
    Dim annee As Long: annee = GCL(cfg, "CFG_Year", Year(Date))
    Dim joursFeries As Object: Set joursFeries = BuildFeriesFast(annee)

    ' === CHARGER DICTIONNAIRES (1 seule fois) ===
    Dim dictCodes As Object: Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    If Not wsCodesSpec Is Nothing Then ChargerSpeciauxFast wsCodesSpec, dictCodes
    If Not wsConfigCodes Is Nothing Then ChargerConfigCodesFast wsConfigCodes, dictCodes, cfg
    
    Dim dictFonctions As Object: Set dictFonctions = ChargerFonctionsFast()
    Dim dictCEFAFormation As Object: Set dictCEFAFormation = ChargerCEFAFormationFast(nomOnglet, couleurINFAdmin)
    
    ' === LECTURE DATA EN BLOC (OPTIMISATION MAJEURE) ===
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Lire TOUTES les donnees en 1 seul acces
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(ligneDebut, 1), ws.Cells(ligneFin, colFin))
    Dim dataArr As Variant: dataArr = dataRange.value
    
    ' Lire TOUTES les couleurs en 1 boucle
    Dim colorArr() As Long
    ReDim colorArr(1 To ligneFin - ligneDebut + 1, 1 To colFin)
    Dim iRow As Long, iCol As Long
    For iRow = ligneDebut To ligneFin
        For iCol = 1 To colFin
            colorArr(iRow - ligneDebut + 1, iCol) = ws.Cells(iRow, iCol).Interior.Color
        Next iCol
    Next iRow
    
    ' Lire numeros de jours
    Dim jourArr As Variant
    jourArr = ws.Range(ws.Cells(ligneNumJour, colDebut), ws.Cells(ligneNumJour, colFin)).value
    
    ' === CLEAR COCKPIT ZONE ===
    Dim rClearStart As Long: rClearStart = rLabelStart
    If rDates < rClearStart Then rClearStart = rDates
    Dim rClearEnd As Long: rClearEnd = rAction
    If rNuitTotal > rClearEnd Then rClearEnd = rNuitTotal
    
    With ws.Range(ws.Cells(rClearStart, colDebut), ws.Cells(rClearEnd + 1, colFin))
        .ClearContents
        .FormatConditions.Delete
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
        .Font.ColorIndex = xlAutomatic
        .Borders.LineStyle = xlNone
    End With
    
    ' Row heights
    If cfgTotalRowHeight > 0 Then
        If lig1 > 0 Then ws.Rows(lig1).RowHeight = cfgTotalRowHeight
        If lig2 > 0 Then ws.Rows(lig2).RowHeight = cfgTotalRowHeight
        If lig3 > 0 Then ws.Rows(lig3).RowHeight = cfgTotalRowHeight
    End If
    
    ' === COPIER DATES AVEC FORMATAGE ===
    If rDatesSrc > 0 Then
        ' Copier valeurs
        ws.Range(ws.Cells(rDates, colDebut), ws.Cells(rDates, colFin)).value = _
            ws.Range(ws.Cells(rDatesSrc, colDebut), ws.Cells(rDatesSrc, colFin)).value
        ' Copier formats (couleurs)
        ws.Range(ws.Cells(rDatesSrc, colDebut), ws.Cells(rDatesSrc, colFin)).Copy
        ws.Range(ws.Cells(rDates, colDebut), ws.Cells(rDates, colFin)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' === BOUCLE PRINCIPALE (OPTIMISEE) ===
    Dim col As Long, i As Long, j As Long
    Dim code As String, nomPersonne As String, fonctionPersonne As String, fctUpper As String
    Dim codeUpper As String, cleNom As String
    Dim couleurCell As Long
    Dim tot(1 To 11) As Double, totINF(1 To 11) As Double, totAS(1 To 11) As Double
    
    ' === COMPTEURS POUR FORMAT "X (Y)" ===
    Dim totEtage(1 To 4) As Double
    Dim totINFSeules(1 To 4) As Double
    
    Dim vLocal(1 To 11) As Double
    Dim c19INF As Long, c19AS As Long, c20Cnt As Long, c20ECnt As Long, c15Cnt As Long
    Dim nuitCode1Cnt As Long, nuitCode2Cnt As Long
    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
    Dim compterDansTotaux As Boolean, codeFound As Boolean
    Dim numJour As Long, dateJour As Date, wd As Integer
    Dim estWE As Boolean, estFerie As Boolean, isValidDay As Boolean
    Dim effPrevu(1 To 4) As Long
    Dim jourRouge As Boolean, delta As Double
    Dim dayTags As String
    Dim isSpecialC20Day As Boolean, isC15Forbidden As Boolean, isC15Required As Boolean
    Dim ok78 As Boolean, total78 As Double, inf78 As Double, as78 As Double
    Dim okC19 As Boolean, c19Total As Long, totalCoupes As Long
    Dim needC15 As Boolean, needC20 As Boolean, needC20E As Boolean, needThirdCoupe As Boolean
    Dim isC19AS As Boolean, isC19INF As Boolean, hasInf81630 As Boolean
    Dim nuitTotal As Long
    Dim arrIdx As Long
    Dim vals As Variant
    
    Dim compteEtage As Boolean
    Dim compteINFSeule As Boolean
    Dim codeNuitNorm As String

    For col = colDebut To colFin
        ' Reset compteurs
        For j = 1 To 11: tot(j) = 0: totINF(j) = 0: totAS(j) = 0: Next j
        For j = 1 To 4: totEtage(j) = 0: totINFSeules(j) = 0: Next j
        c19INF = 0: c19AS = 0: c20Cnt = 0: c20ECnt = 0: c15Cnt = 0
        nuitCode1Cnt = 0: nuitCode2Cnt = 0

        ' === INNER LOOP (lecture depuis arrays) ===
        For i = ligneDebut To ligneFin
            arrIdx = i - ligneDebut + 1
            couleurCell = colorArr(arrIdx, col)
            
            ' Rouge = absence, toujours ignorer
            If couleurCell = 255 Then GoTo NextPerson
            
            If couleurCell <> couleurIgnore Then
                code = Trim$(CStr(dataArr(arrIdx, col)))
                
                If Len(code) > 0 Then
                    nomPersonne = Trim$(CStr(dataArr(arrIdx, 1)))
                    
                    ' Fonction lookup
                    fonctionPersonne = ""
                    If dictFonctions.Exists(nomPersonne) Then
                        fonctionPersonne = dictFonctions(nomPersonne)
                    Else
                        cleNom = Replace$(nomPersonne, " ", "_")
                        If dictFonctions.Exists(cleNom) Then
                            fonctionPersonne = dictFonctions(cleNom)
                        End If
                    End If
                    fctUpper = UCase$(fonctionPersonne)
                    
                    ' Check fonction autorisee
                    compterDansTotaux = (InStr(fonctionsACompter, "," & fctUpper & ",") > 0)
                    
                    ' CEFA en formation?
                    If compterDansTotaux And fctUpper = "CEFA" Then
                        If dictCEFAFormation.Exists(cleNom) Then compterDansTotaux = False
                    End If
                    
                    ' Exclusion codes absence
                    If compterDansTotaux Then
                        codeUpper = UCase$(code)
                        ' Supprimer apostrophe de debut si presente
                        If Left$(codeUpper, 1) = "'" Then codeUpper = Mid$(codeUpper, 2)
                        
                        If codeUpper = "WE" Or Left$(codeUpper, 3) = "MAL" Or Left$(codeUpper, 2) = "CA" Or _
                           Left$(codeUpper, 3) = "RCT" Or Left$(codeUpper, 3) = "MAT" Or Left$(codeUpper, 3) = "MUT" Or _
                           codeUpper = "CTR" Or codeUpper = "DP" Or codeUpper = "RHS" Or codeUpper = "EL" Or _
                           Left$(codeUpper, 3) = "AFC" Or Left$(codeUpper, 3) = "3/4" Or Left$(codeUpper, 3) = "4/5" Or _
                           Left$(codeUpper, 2) = "CP" Or Left$(codeUpper, 1) = "R" Or Left$(codeUpper, 1) = "F" Then
                            compterDansTotaux = False
                        End If
                    End If
                    
                    ' === OBTENIR VALEURS ===
                    For j = 1 To 11: vLocal(j) = 0: Next j
                    codeFound = False
                    
                    Dim codeNorm As String: codeNorm = UCase$(Replace$(code, " ", ""))
                    ' Supprimer apostrophe
                    If Left$(codeNorm, 1) = "'" Then codeNorm = Mid$(codeNorm, 2)
                    
                    ' Try exact match first
                    If dictCodes.Exists(code) Then
                        vals = dictCodes(code)
                        For j = 1 To 11: vLocal(j) = vals(j): Next j
                        codeFound = True
                    Else
                        ' Try prefix match
                        Dim foundByPrefix As Boolean: foundByPrefix = False
                        Dim dictKey As Variant
                        For Each dictKey In dictCodes.keys
                            If Left$(CStr(dictKey), Len(code)) = code Or _
                               Left$(CStr(dictKey), Len(code) + 1) = code & " " Then
                                vals = dictCodes(dictKey)
                                For j = 1 To 11: vLocal(j) = vals(j): Next j
                                codeFound = True
                                foundByPrefix = True
                                Exit For
                            End If
                        Next dictKey
                        
                        If Not foundByPrefix Then
                            If ParseCodeFast(code, h1, f1, h2, f2) Then
                                CalcPeriodesFast h1, f1, h2, f2, vLocal(1), vLocal(2), vLocal(3), vLocal(4)
                                CalcPresSpecFast h1, f1, h2, f2, vLocal(5), vLocal(6), vLocal(7)
                                DetectSpecialCodes h1, f1, h2, f2, vLocal(8), vLocal(9), vLocal(10), vLocal(11)
                                codeFound = True
                            End If
                        End If
                    End If
                    
                    ' Detect C codes
                    If codeNorm Like "C20E*" Or codeNorm Like "*C20E*" Then
                        vLocal(10) = 1: vLocal(9) = 0: vLocal(8) = 0
                        vLocal(1) = 1: vLocal(2) = 0: vLocal(3) = 1: vLocal(4) = 0
                        codeFound = True
                    ElseIf codeNorm Like "C20*" Or codeNorm Like "*C20" Then
                        If Not (codeNorm Like "C20E*" Or codeNorm Like "*C20E*") Then
                            vLocal(9) = 1: vLocal(10) = 0: vLocal(8) = 0
                            vLocal(1) = 1: vLocal(2) = 0: vLocal(3) = 1: vLocal(4) = 0
                            codeFound = True
                        End If
                    ElseIf codeNorm Like "C15*" Or codeNorm Like "*C15*" Then
                        vLocal(8) = 1: vLocal(9) = 0: vLocal(10) = 0
                        vLocal(1) = 1: vLocal(2) = 0: vLocal(3) = 1: vLocal(4) = 0
                        codeFound = True
                    ElseIf codeNorm Like "C19*" Or codeNorm Like "*C19*" Then
                        vLocal(11) = 1
                        vLocal(1) = 1: vLocal(2) = 0: vLocal(3) = 1: vLocal(4) = 0
                        codeFound = True
                    End If
                    
                    If codeFound And compterDansTotaux Then
                        For j = 1 To 11: tot(j) = tot(j) + vLocal(j): Next j
                        
                        ' === COMPTAGE SEPARE POUR FORMAT "X (Y)" ===
                        compteEtage = True
                        compteINFSeule = False
                        
                        Select Case fctUpper
                            Case "INF"
                                If couleurCell = couleurINFAdmin Then
                                    compteEtage = False
                                Else
                                    compteINFSeule = True
                                End If
                            Case "AS"
                                If couleurCell = couleurBleuClair Then
                                    compteEtage = False
                                End If
                            Case "CEFA"
                                compteEtage = True
                        End Select
                        
                        If compteEtage Then
                            For j = 1 To 4: totEtage(j) = totEtage(j) + vLocal(j): Next j
                        End If
                        If compteINFSeule Then
                            For j = 1 To 4: totINFSeules(j) = totINFSeules(j) + vLocal(j): Next j
                        End If
                        
                        ' Comptage speciaux
                        If vLocal(8) > 0 Then c15Cnt = c15Cnt + 1
                        If vLocal(9) > 0 Then c20Cnt = c20Cnt + 1
                        If vLocal(10) > 0 Then c20ECnt = c20ECnt + 1
                        If vLocal(11) > 0 Then
                            If fctUpper = "INF" Then c19INF = c19INF + 1 Else c19AS = c19AS + 1
                        End If
                        
                        ' === MATCHING CODES NUIT (FIX v5 - apostrophe + normalisation) ===
                        codeUpper = UCase$(Trim$(code))
                        ' Supprimer apostrophe de debut
                        If Left$(codeUpper, 1) = "'" Then codeUpper = Mid$(codeUpper, 2)
                        ' Normaliser pour comparaison
                        codeNuitNorm = NormaliserCodeNuit(codeUpper)
                        
                        If codeNuitNorm = cfgNuitCode1Norm Then
                            nuitCode1Cnt = nuitCode1Cnt + 1
                        End If
                        If codeNuitNorm = cfgNuitCode2Norm Then
                            nuitCode2Cnt = nuitCode2Cnt + 1
                        End If
                        
                        ' Comptage INF/AS global
                        If fctUpper = "INF" And couleurCell <> couleurINFAdmin Then
                            For j = 1 To 11: totINF(j) = totINF(j) + vLocal(j): Next j
                        End If
                        If fctUpper = "AS" Then
                            For j = 1 To 11: totAS(j) = totAS(j) + vLocal(j): Next j
                        End If
                    End If
                End If
            End If
NextPerson:
        Next i

        ' === COCKPIT RENDER ===
        numJour = 0
        On Error Resume Next
        numJour = CLng(jourArr(1, col - colDebut + 1))
        On Error GoTo 0
        
        isValidDay = (numJour > 0 And numJour <= 31)
        If isValidDay Then
            dateJour = DateSerial(annee, MoisNumFast(nomOnglet), numJour)
            wd = Weekday(dateJour, vbMonday)
            estWE = (wd >= 6)
            estFerie = joursFeries.Exists(CStr(dateJour))
        Else
            estWE = False: estFerie = False: wd = 0
        End If
        
        ' Effectifs prevus
        If estFerie Then
            effPrevu(1) = effFER(1): effPrevu(2) = effFER(2): effPrevu(3) = effFER(3): effPrevu(4) = effFER(4)
        ElseIf estWE Then
            effPrevu(1) = effWE(1): effPrevu(2) = effWE(2): effPrevu(3) = effWE(3): effPrevu(4) = effWE(4)
        Else
            effPrevu(1) = effSem(1): effPrevu(2) = effSem(2): effPrevu(3) = effSem(3): effPrevu(4) = effSem(4)
        End If
        
        ' === ECRITURE COCKPIT ===
        jourRouge = False
        
        ' === TOTAUX LIGNES 60-62 - FORMAT "X (Y)" ===
        If lig1 > 0 Then WriteEtageINFCell ws.Cells(lig1, col), totEtage(1), totINFSeules(1), effPrevu(1), isValidDay, cfgTotalFontName, cfgTotalFontSize, cfgTotalOkColor, cfgTotalBadColor, cfgTotalOkFontColor, cfgTotalBadFontColor
        If lig2 > 0 Then WriteEtageINFCell ws.Cells(lig2, col), totEtage(2), totINFSeules(2), effPrevu(2), isValidDay, cfgTotalFontName, cfgTotalFontSize, cfgTotalOkColor, cfgTotalBadColor, cfgTotalOkFontColor, cfgTotalBadFontColor
        If lig3 > 0 Then WriteEtageINFCell ws.Cells(lig3, col), totEtage(3), totINFSeules(3), effPrevu(3), isValidDay, cfgTotalFontName, cfgTotalFontSize, cfgTotalOkColor, cfgTotalBadColor, cfgTotalOkFontColor, cfgTotalBadFontColor
        
        ' Delta rows
        delta = totEtage(1) - effPrevu(1)
        DrawDeltaFast ws.Cells(rMatin, col), delta, False, cfgDeltaFontName, cfgDeltaFontSize, totEtage(1)
        If delta < 0 Then jourRouge = True
        
        delta = totEtage(2) - effPrevu(2)
        DrawDeltaFast ws.Cells(rAM, col), delta, False, cfgDeltaFontName, cfgDeltaFontSize, totEtage(2)
        If delta < 0 Then jourRouge = True
        
        delta = totEtage(3) - effPrevu(3)
        DrawDeltaFast ws.Cells(rSoir, col), delta, False, cfgDeltaFontName, cfgDeltaFontSize, totEtage(3)
        If delta < 0 Then jourRouge = True
        
        delta = totEtage(4) - effPrevu(4)
        DrawDeltaFast ws.Cells(rNuit, col), delta, False, cfgDeltaFontName, cfgDeltaFontSize, totEtage(4)
        If delta < 0 Then jourRouge = True
        
        ' 7h-8h critique
        total78 = tot(6): inf78 = totINF(6): as78 = totAS(6)
        If total78 < cfg78MinTotal Then
            ok78 = False: delta = total78 - cfg78MinTotal
        ElseIf cfg78AllowAllInf <> 0 And inf78 >= cfg78MinTotal Then
            ok78 = True: delta = 0
        Else
            ok78 = (inf78 >= cfg78MinInf And as78 >= cfg78MinAS)
            delta = IIf(ok78, 0, -1)
        End If
        DrawDeltaFast ws.Cells(r7h8h, col), delta, True, cfgDeltaFontName, cfgDeltaFontSize, 0
        If Not ok78 Then jourRouge = True
        
        ' Presences speciales
        WriteCheckCell ws.Cells(rP0645, col), tot(5), cfgP0645Req, cfgCheckFontName, cfgCheckFontOk, cfgCheckFontAlert, jourRouge
        WriteValueCell ws.Cells(rP7H8H, col), tot(6), cfgCheckFontName, cfgCheckFontOk
        WriteValueCell ws.Cells(rP81630, col), tot(7), cfgCheckFontName, cfgCheckFontOk
        
        ' Couvertures C15/C20/C20E/C19
        dayTags = GetDayTagsFast(wd, estFerie)
        isSpecialC20Day = DayInListFast(dayTags, cfgSpecialC20)
        isC15Forbidden = DayInListFast(dayTags, cfgC15Forbidden)
        isC15Required = DayInListFast(dayTags, cfgC15Required)
        
        c19Total = c19INF + c19AS
        isC19AS = (c19AS > 0 And c19INF = 0)
        isC19INF = (c19INF > 0)
        hasInf81630 = (totINF(7) > 0)
        totalCoupes = c15Cnt + c20Cnt + c20ECnt + c19INF + c19AS
        
        needC15 = False: needC20 = False: needC20E = False: needThirdCoupe = False
        If isSpecialC20Day Then
            If isC19AS And Not hasInf81630 Then needC20E = True
        Else
            If isC15Required And isC19AS Then needC15 = True
            If isC19INF Then needC20 = True
            If isC19AS And Not hasInf81630 Then needC20E = True
            If isC19INF And totalCoupes < cfgCoupeMinTotal And (c20Cnt + c20ECnt) >= 1 Then needThirdCoupe = True
        End If
        
        WriteValueCell ws.Cells(rC15, col), CDbl(c15Cnt), cfgCheckFontName, cfgCheckFontOk
        WriteValueCell ws.Cells(rC20, col), CDbl(c20Cnt), cfgCheckFontName, cfgCheckFontOk
        WriteValueCell ws.Cells(rC20E, col), CDbl(c20ECnt), cfgCheckFontName, cfgCheckFontOk
        WriteValueCell ws.Cells(rC19, col), CDbl(c19Total), cfgCheckFontName, cfgCheckFontOk
        
        okC19 = (c19Total >= cfgC19Req)
        If Not okC19 Then
            WriteAlertCell ws.Cells(rC19, col), "MANQUE", cfgCheckFontName, cfgCheckFontAlert
            jourRouge = True
        End If
        
        If isSpecialC20Day Then
            If c20Cnt + c20ECnt < cfgC20ReqSpecial Then
                WriteAlertCell ws.Cells(rC20, col), "MANQUE C20", cfgCheckFontName, cfgCheckFontAlert
                jourRouge = True
            End If
            If needC20E And c20ECnt < 1 Then
                WriteAlertCell ws.Cells(rC20E, col), "MANQUE C20E", cfgCheckFontName, cfgCheckFontAlert
                jourRouge = True
            End If
            If isC15Forbidden And c15Cnt > 0 Then
                WriteAlertCell ws.Cells(rC15, col), "C15 INTERDIT", cfgCheckFontName, cfgCheckFontAlert
                jourRouge = True
            End If
        Else
            If needC20 And c20Cnt < 1 Then
                WriteAlertCell ws.Cells(rC20, col), "MANQUE C20", cfgCheckFontName, cfgCheckFontAlert
                jourRouge = True
            End If
            If needC20E And c20ECnt < 1 Then
                WriteAlertCell ws.Cells(rC20E, col), "MANQUE C20E", cfgCheckFontName, cfgCheckFontAlert
                jourRouge = True
            End If
            If needC15 And c15Cnt < cfgC15ReqCount Then
                WriteAlertCell ws.Cells(rC15, col), "MANQUE", cfgCheckFontName, cfgCheckFontAlert
                jourRouge = True
            End If
            If needThirdCoupe And c15Cnt = 0 And c20ECnt = 0 Then
                WriteAlertCell ws.Cells(rC15, col), "MANQUE", cfgCheckFontName, cfgCheckFontAlert
                jourRouge = True
            End If
        End If
        
        ' === NUIT CODES (lignes 74-76) ===
        nuitTotal = nuitCode1Cnt + nuitCode2Cnt
        WriteValueCell ws.Cells(rNuitCode1, col), CDbl(nuitCode1Cnt), cfgCheckFontName, cfgCheckFontOk
        WriteValueCell ws.Cells(rNuitCode2, col), CDbl(nuitCode2Cnt), cfgCheckFontName, cfgCheckFontOk
        WriteValueCell ws.Cells(rNuitTotal, col), CDbl(nuitTotal), cfgCheckFontName, cfgCheckFontOk
        If cfgNuitReqTotal > 0 And nuitTotal < cfgNuitReqTotal Then
            With ws.Cells(rNuitTotal, col)
                .Interior.Color = RGB(255, 200, 150)
                .Font.Size = cfgCheckFontAlert
            End With
            jourRouge = True
        End If
        
        ' Meteo + Action
        With ws.Cells(rMeteo, col)
            If jourRouge Then .value = ChrW(&H26A0): .Font.Color = vbRed Else .value = ""
            .HorizontalAlignment = xlCenter
        End With
        
        If jourRouge Then
            With ws.Cells(rAction, col)
                .value = "APPEL"
                .Interior.Color = vbBlack
                .Font.Color = vbWhite
                .Font.Bold = True
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
            End With
        End If
    Next col

    ' === LABELS COCKPIT ===
    ApplyLabelsFast ws, lig1, lig2, lig3, rMeteo, rMatin, r7h8h, rAM, rSoir, rNuit, _
                    rP0645, rP7H8H, rP81630, rC15, rC20, rC20E, rC19, _
                    rNuitCode1, rNuitCode2, rNuitTotal, rAction, _
                    cfgNuitCode1, cfgNuitCode2, _
                    rLabelStart, rLabelEnd, rDates, rDatesSrc, colDebut, colFin, _
                    cfgLabelFontName, cfgLabelFontSize, cfgLabelFontBold, _
                    cfgTotalLabelFontSize, cfgTotalLabelFontColor

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
' === RESTAURER MASQUAGE SELON MODE ACTIF (Jour ou Nuit) ===
    ' Detection du mode:
    ' - Si lignes 31-38 (nuit) sont masquees = Mode Jour actif
    ' - Si lignes 6-28 (jour) sont masquees = Mode Nuit actif
    Dim modeJourActif As Boolean: modeJourActif = False
    Dim modeNuitActif As Boolean: modeNuitActif = False

    ' Lire les blocs de masquage depuis config
    Dim hideBlocksJour As String: hideBlocksJour = GCS(cfg, "VIEW_Jour_HideBlocks")
    Dim hideBlocksNuit As String: hideBlocksNuit = GCS(cfg, "VIEW_Nuit_HideBlocks")

    ' Detecter le mode actuel (AVANT que le calcul n'ait tout modifie)
    ' Si ligne 31 (personnel nuit) est masquee, on est en Mode Jour
    ' Si ligne 6 (personnel jour) est masquee, on est en Mode Nuit
    If ws.Rows(31).Hidden Then
        modeJourActif = True
    ElseIf ws.Rows(6).Hidden Then
        modeNuitActif = True
    End If

    Dim blocks() As String, blk As Variant
    Dim rngParts() As String, startRow As Long, endRow As Long

    ' Appliquer masquage Mode Jour si actif
    If modeJourActif And Len(hideBlocksJour) > 0 Then
        blocks = Split(hideBlocksJour, ";")
        For Each blk In blocks
            If InStr(blk, ":") > 0 Then
                rngParts = Split(blk, ":")
                On Error Resume Next
                startRow = CLng(Trim$(rngParts(0)))
                endRow = CLng(Trim$(rngParts(1)))
                On Error GoTo 0
                If startRow > 0 And endRow >= startRow Then
                    ws.Rows(startRow & ":" & endRow).Hidden = True
                End If
            End If
        Next blk
        ' Masquer colonne B en mode Jour
        ws.Columns("B").Hidden = True
    End If

    ' Appliquer masquage Mode Nuit si actif
    If modeNuitActif And Len(hideBlocksNuit) > 0 Then
        blocks = Split(hideBlocksNuit, ";")
        For Each blk In blocks
            If InStr(blk, ":") > 0 Then
                rngParts = Split(blk, ":")
                On Error Resume Next
                startRow = CLng(Trim$(rngParts(0)))
                endRow = CLng(Trim$(rngParts(1)))
                On Error GoTo 0
                If startRow > 0 And endRow >= startRow Then
                    ws.Rows(startRow & ":" & endRow).Hidden = True
                End If
            End If
        Next blk
        ' Masquer colonne B en mode Nuit aussi
        ws.Columns("B").Hidden = True
    End If

    MsgBox "Cockpit OK - " & Format$(Timer - startTime, "0.00") & "s", vbInformation
End Sub

' =============================================================================
' FONCTION NORMALISATION CODE NUIT (FIX v5)
' Supprime apostrophe, normalise espaces, uniformise format heure
' =============================================================================
Private Function NormaliserCodeNuit(s As String) As String
    Dim tmp As String: tmp = Trim$(s)
    ' Supprimer apostrophe de debut
    If Left$(tmp, 1) = "'" Then tmp = Mid$(tmp, 2)
    ' Supprimer espaces multiples
    Do While InStr(tmp, "  ") > 0: tmp = Replace$(tmp, "  ", " "): Loop
    ' Uniformiser H et h en :
    tmp = Replace$(tmp, "H", ":")
    tmp = Replace$(tmp, "h", ":")
    ' Mettre en majuscules
    NormaliserCodeNuit = UCase$(tmp)
End Function

' =============================================================================
' FONCTIONS HELPER
' =============================================================================

Private Function IsMonthSheet(n As String) As Boolean
    IsMonthSheet = (n = "Janv" Or n = "Fev" Or n = "Mars" Or n = "Avril" Or n = "Mai" Or _
                    n = "Juin" Or n = "Juil" Or n = "Aout" Or n = "Sept" Or n = "Oct" Or _
                    n = "Nov" Or n = "Dec")
End Function

Private Function MoisNumFast(n As String) As Integer
    Select Case Left$(n, 3)
        Case "Jan": MoisNumFast = 1
        Case "Fev": MoisNumFast = 2
        Case "Mar": MoisNumFast = 3
        Case "Avr": MoisNumFast = 4
        Case "Mai": MoisNumFast = 5
        Case "Jui": MoisNumFast = IIf(Left$(n, 4) = "Juil", 7, 6)
        Case "Aou": MoisNumFast = 8
        Case "Sep": MoisNumFast = 9
        Case "Oct": MoisNumFast = 10
        Case "Nov": MoisNumFast = 11
        Case "Dec": MoisNumFast = 12
        Case Else: MoisNumFast = 1
    End Select
End Function

Private Function ChargerConfigFast(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If ws Is Nothing Then Set ChargerConfigFast = d: Exit Function
    Dim lr As Long: lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Set ChargerConfigFast = d: Exit Function
    Dim arr As Variant: arr = ws.Range("A2:B" & lr).value
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        If Len(Trim$(CStr(arr(i, 1)))) > 0 Then
            If Not d.Exists(arr(i, 1)) Then d(Trim$(CStr(arr(i, 1)))) = arr(i, 2)
        End If
    Next i
    Set ChargerConfigFast = d
End Function

Private Function GCL(d As Object, k As String, def As Long) As Long
    If d.Exists(k) Then If IsNumeric(d(k)) Then GCL = CLng(d(k)): Exit Function
    GCL = def
End Function

Private Function GCS(d As Object, k As String) As String
    If d.Exists(k) Then GCS = CStr(d(k)) Else GCS = ""
End Function

Private Function BuildFeriesFast(annee As Long) As Object
    Dim f As Object: Set f = CreateObject("Scripting.Dictionary")
    Dim p As Date: p = CalculerPaquesFast(annee)
    On Error Resume Next
    f.Add CStr(DateSerial(annee, 1, 1)), 1
    f.Add CStr(p + 1), 1
    f.Add CStr(DateSerial(annee, 5, 1)), 1
    f.Add CStr(p + 39), 1
    f.Add CStr(p + 50), 1
    f.Add CStr(DateSerial(annee, 7, 21)), 1
    f.Add CStr(DateSerial(annee, 8, 15)), 1
    f.Add CStr(DateSerial(annee, 11, 1)), 1
    f.Add CStr(DateSerial(annee, 11, 11)), 1
    f.Add CStr(DateSerial(annee, 12, 25)), 1
    On Error GoTo 0
    Set BuildFeriesFast = f
End Function

Private Function CalculerPaquesFast(annee As Long) As Date
    Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
    Dim f As Integer, g As Integer, h As Integer, i As Integer, k As Integer
    Dim l As Integer, m As Integer, mois As Integer, jour As Integer
    a = annee Mod 19: b = annee \ 100: c = annee Mod 100
    d = b \ 4: e = b Mod 4: f = (b + 8) \ 25: g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30: i = c \ 4: k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7: m = (a + 11 * h + 22 * l) \ 451
    mois = (h + l - 7 * m + 114) \ 31: jour = ((h + l - 7 * m + 114) Mod 31) + 1
    CalculerPaquesFast = DateSerial(annee, mois, jour)
End Function

Private Function ChargerFonctionsFast() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("Personnel"): On Error GoTo 0
    If ws Is Nothing Then Set ChargerFonctionsFast = d: Exit Function
    Dim lr As Long: lr = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lr < 2 Then Set ChargerFonctionsFast = d: Exit Function
    Dim arr As Variant: arr = ws.Range("B2:E" & lr).value
    Dim i As Long, nom As String, prenom As String, fct As String
    Dim cleUnder As String
    For i = 1 To UBound(arr, 1)
        nom = Trim$(CStr(arr(i, 1)))
        prenom = Trim$(CStr(arr(i, 2)))
        fct = UCase$(Trim$(CStr(arr(i, 4))))
        If nom <> "" And prenom <> "" Then
            cleUnder = nom & "_" & prenom
            If Not d.Exists(cleUnder) Then d.Add cleUnder, fct
        End If
    Next i
    Set ChargerFonctionsFast = d
End Function

Private Function ChargerCEFAFormationFast(nomMois As String, couleurFormation As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("Personnel"): On Error GoTo 0
    If ws Is Nothing Then Set ChargerCEFAFormationFast = d: Exit Function
    Dim colMoisPct As Long: colMoisPct = 29 + (MoisNumFast(nomMois) - 1) * 2
    Dim lr As Long: lr = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lr < 2 Then Set ChargerCEFAFormationFast = d: Exit Function
    Dim i As Long, cle As String, fct As String
    For i = 2 To lr
        fct = UCase$(Trim$(CStr(ws.Cells(i, "E").value)))
        If fct = "CEFA" Then
            cle = Trim$(CStr(ws.Cells(i, "B").value)) & "_" & Trim$(CStr(ws.Cells(i, "C").value))
            If ws.Cells(i, colMoisPct).Interior.Color = couleurFormation Then d(cle) = True
        End If
    Next i
    Set ChargerCEFAFormationFast = d
End Function

Private Sub ChargerSpeciauxFast(ws As Worksheet, d As Object)
    Dim lr As Long: lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Exit Sub
    Dim arr As Variant: arr = ws.Range("A2:G" & lr).value
    Dim i As Long, k As Long, code As String, cc As String, v(1 To 11) As Double
    For i = 1 To UBound(arr, 1)
        code = Trim$(CStr(arr(i, 1)))
        If Len(code) > 0 And Not d.Exists(code) Then
            For k = 1 To 11: v(k) = 0: Next k
            If IsNumeric(arr(i, 2)) Then v(5) = CDbl(arr(i, 2))
            If IsNumeric(arr(i, 3)) Then v(6) = CDbl(arr(i, 3))
            If IsNumeric(arr(i, 4)) Then v(1) = CDbl(arr(i, 4))
            If IsNumeric(arr(i, 5)) Then v(2) = CDbl(arr(i, 5))
            If IsNumeric(arr(i, 6)) Then v(3) = CDbl(arr(i, 6))
            If IsNumeric(arr(i, 7)) Then v(4) = CDbl(arr(i, 7))
            cc = UCase$(Replace$(code, " ", ""))
            If cc Like "C19*" Then v(11) = 1: If v(6) = 0 Then v(6) = 1
            If cc Like "C20E*" Then v(10) = 1
            If cc Like "C20*" And Not cc Like "C20E*" Then v(9) = 1
            If cc Like "C15*" Then v(8) = 1
            d.Add code, v
        End If
    Next i
End Sub

Private Sub ChargerConfigCodesFast(ws As Worksheet, d As Object, cfg As Object)
    Dim lr As Long: lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Exit Sub
    Dim hasExt As Boolean: hasExt = (ws.Cells(1, 6).value <> "")
    Dim arr As Variant
    If hasExt Then arr = ws.Range("A2:O" & lr).value Else arr = ws.Range("A2:A" & lr).value
    Dim i As Long, k As Long, code As String, v(1 To 11) As Double
    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
    For i = 1 To UBound(arr, 1)
        code = Trim$(CStr(arr(i, 1)))
        If Len(code) > 0 And Not d.Exists(code) Then
            For k = 1 To 11: v(k) = 0: Next k
            If hasExt And UBound(arr, 2) >= 15 Then
                If IsNumeric(arr(i, 12)) And arr(i, 12) <> "" Then v(1) = CDbl(arr(i, 12))
                If IsNumeric(arr(i, 13)) And arr(i, 13) <> "" Then v(2) = CDbl(arr(i, 13))
                If IsNumeric(arr(i, 14)) And arr(i, 14) <> "" Then v(3) = CDbl(arr(i, 14))
                If IsNumeric(arr(i, 15)) And arr(i, 15) <> "" Then v(4) = CDbl(arr(i, 15))
                If IsNumeric(arr(i, 10)) And arr(i, 10) <> "" Then v(5) = CDbl(arr(i, 10))
                If IsNumeric(arr(i, 11)) And arr(i, 11) <> "" Then v(6) = CDbl(arr(i, 11))
            End If
            If v(1) = 0 And v(2) = 0 And v(3) = 0 And v(4) = 0 Then
                If ParseCodeFast(code, h1, f1, h2, f2) Then
                    CalcPeriodesFast h1, f1, h2, f2, v(1), v(2), v(3), v(4)
                    CalcPresSpecFast h1, f1, h2, f2, v(5), v(6), v(7)
                    DetectSpecialCodes h1, f1, h2, f2, v(8), v(9), v(10), v(11)
                End If
            End If
            d.Add code, v
        End If
    Next i
End Sub

Private Function ParseCodeFast(c As String, ByRef s1 As Double, ByRef e1 As Double, ByRef s2 As Double, ByRef e2 As Double) As Boolean
    s1 = 0: e1 = 0: s2 = 0: e2 = 0
    Dim tmp As String: tmp = Trim$(Replace$(Replace$(c, vbLf, " "), vbCr, " "))
    ' Supprimer apostrophe
    If Left$(tmp, 1) = "'" Then tmp = Mid$(tmp, 2)
    Do While InStr(tmp, "  ") > 0: tmp = Replace$(tmp, "  ", " "): Loop
    Dim p() As String: p = Split(tmp, " ")
    On Error GoTo Err1
    If UBound(p) = 1 Then
        s1 = HDecFast(p(0)): e1 = HDecFast(p(1))
        If s1 > 0 Or e1 > 0 Then ParseCodeFast = True
    ElseIf UBound(p) >= 3 Then
        s1 = HDecFast(p(0)): e1 = HDecFast(p(1)): s2 = HDecFast(p(2)): e2 = HDecFast(p(3))
        If s1 > 0 Or e1 > 0 Or s2 > 0 Or e2 > 0 Then ParseCodeFast = True
    End If
    Exit Function
Err1: ParseCodeFast = False
End Function

Private Function HDecFast(s As String) As Double
    If Len(s) = 0 Then HDecFast = 0: Exit Function
    If InStr(s, ":") > 0 Then
        Dim p() As String: p = Split(s, ":")
        On Error Resume Next: HDecFast = CDbl(p(0)) + CDbl(p(1)) / 60: On Error GoTo 0
    ElseIf IsNumeric(s) Then
        Dim v As Double: v = CDbl(s)
        If v < 1 And v > 0 Then HDecFast = v * 24 Else HDecFast = v
    End If
End Function

Private Sub CalcPeriodesFast(h1 As Double, f1 As Double, h2 As Double, f2 As Double, ByRef mat As Double, ByRef am As Double, ByRef soi As Double, ByRef nui As Double)
    mat = 0: am = 0: soi = 0: nui = 0
    Dim fin As Double: fin = IIf(f2 > 0, f2, f1)
    If h1 = 0 And f1 = 0 Then Exit Sub
    If h1 < SEUIL_MATIN Or (h2 > 0 And h2 < SEUIL_MATIN) Then mat = 1
    If fin > SEUIL_PM Or (h2 > 0 And f2 > SEUIL_PM) Then am = 1
    If fin > SEUIL_SOIR_PLEIN Then
        soi = 1
    ElseIf fin > SEUIL_SOIR_DEMI Then
        soi = 0.5
    End If
    If h1 >= SEUIL_NUIT_DEBUT Or (fin > 0 And fin <= SEUIL_NUIT_FIN) Then
        nui = IIf(Abs(fin - 24) < 0.1 Or fin = 0, 0.5, 1)
    End If
End Sub

Private Sub CalcPresSpecFast(h1 As Double, f1 As Double, h2 As Double, f2 As Double, ByRef p645 As Double, ByRef p78 As Double, ByRef p1630 As Double)
    p645 = IIf(h1 <= 6.75, 1, 0)
    p78 = IIf(h1 < 8 And f1 > 7, 1, 0)
    p1630 = IIf(Abs(f1 - 16.5) < 0.25 Or Abs(f2 - 16.5) < 0.25, 1, 0)
End Sub

Private Sub DetectSpecialCodes(h1 As Double, f1 As Double, h2 As Double, f2 As Double, ByRef c15 As Double, ByRef c20 As Double, ByRef c20e As Double, ByRef c19 As Double)
    c15 = 0: c20 = 0: c20e = 0: c19 = 0
    Dim fin As Double: fin = IIf(f2 > 0, f2, f1)
    If fin >= 15 And fin <= 15.5 Then c15 = 1
    If fin >= 19.75 And fin <= 20.25 Then c20 = 1
    If fin > 20.25 And fin <= 21 Then c20e = 1
    If fin >= 18.75 And fin <= 19.25 Then c19 = 1
End Sub

Private Sub DrawDeltaFast(rng As Range, delta As Double, isCritical As Boolean, fontName As String, fontSize As Long, totVal As Double)
    With rng
        .NumberFormat = "General"
        .Font.Name = fontName
        .Font.Size = fontSize
        .HorizontalAlignment = xlCenter
        If delta < 0 Then
            If isCritical Then
                .value = "!" & delta
                .Interior.Color = RGB(255, 0, 0)
                .Font.Color = vbWhite
            Else
                .value = delta
                .Interior.Color = RGB(255, 200, 200)
                .Font.Color = RGB(200, 0, 0)
            End If
            .Font.Bold = True
        Else
            .value = IIf(totVal > 0, totVal, "")
            .Font.Bold = False
            .Font.Color = RGB(160, 160, 160)
            .Interior.ColorIndex = xlNone
        End If
    End With
End Sub

Private Sub WriteEtageINFCell(rng As Range, totalEtage As Double, totalINF As Double, _
                              expected As Long, isValidDay As Boolean, _
                              fontName As String, fontSize As Long, _
                              okColor As Long, badColor As Long, _
                              okFontColor As Long, badFontColor As Long)
    With rng
        .NumberFormat = "@"
        If Not isValidDay Or (totalEtage = 0 And expected = 0) Then
            .value = ""
            .Font.Bold = False
            .Font.Color = vbBlack
            .Interior.ColorIndex = xlNone
            Exit Sub
        End If
        
        .value = Format$(totalEtage, "0") & " (" & Format$(totalINF, "0") & ")"
        .HorizontalAlignment = xlCenter
        .Font.Name = fontName
        .Font.Size = fontSize
        
        If totalEtage < expected Then
            .Interior.Color = badColor
            .Font.Color = badFontColor
            .Font.Bold = True
        Else
            .Interior.Color = okColor
            .Font.Color = okFontColor
            .Font.Bold = False
        End If
    End With
End Sub

Private Sub WriteValueCell(rng As Range, val As Double, fontName As String, fontSize As Long)
    With rng
        .NumberFormat = "General"
        .Font.Name = fontName
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        If val > 0 Then
            .value = val
        Else
            .value = ""
        End If
    End With
End Sub

Private Sub WriteCheckCell(rng As Range, val As Double, req As Long, fontName As String, fontOk As Long, fontAlert As Long, ByRef jourRouge As Boolean)
    With rng
        .NumberFormat = "General"
        .Font.Name = fontName
        .HorizontalAlignment = xlCenter
        If val > 0 Then
            .value = val
            .Font.Size = 16
        Else
            .value = ""
            .Font.Size = 16
        End If
        If val < req Then
            .Interior.Color = RGB(255, 200, 150)
            If val = 0 Then
                .value = "MANQUE"
                .Font.Size = 9
            End If
            jourRouge = True
        End If
    End With
End Sub

Private Sub WriteAlertCell(rng As Range, txt As String, fontName As String, fontSize As Long)
    With rng
        .value = txt
        .Interior.Color = RGB(255, 200, 150)
        .Font.Size = 9
        .Font.Name = fontName
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Function GetDayTagsFast(wd As Integer, estFerie As Boolean) As String
    Dim t As String
    Select Case wd
        Case 1: t = "LUN": Case 2: t = "MAR": Case 3: t = "MER": Case 4: t = "JEU"
        Case 5: t = "VEN": Case 6: t = "SAM": Case 7: t = "DIM": Case Else: t = ""
    End Select
    If estFerie Then t = t & ",FERIE"
    GetDayTagsFast = t
End Function

Private Function DayInListFast(tags As String, listStr As String) As Boolean
    DayInListFast = False
    If Len(listStr) = 0 Then Exit Function
    Dim tagsU As String: tagsU = "," & UCase$(tags) & ","
    Dim arr() As String: arr = Split(UCase$(Replace$(listStr, ";", ",")), ",")
    Dim i As Long, tok As String
    For i = 0 To UBound(arr)
        tok = Trim$(arr(i))
        If Len(tok) > 0 Then
            If tok = "WE" Then
                If InStr(tagsU, ",SAM,") > 0 Or InStr(tagsU, ",DIM,") > 0 Then DayInListFast = True: Exit Function
            ElseIf InStr(tagsU, "," & tok & ",") > 0 Then
                DayInListFast = True: Exit Function
            End If
        End If
    Next i
End Function

Private Sub ApplyLabelsFast(ws As Worksheet, lig1 As Long, lig2 As Long, lig3 As Long, _
    rMeteo As Long, rMatin As Long, r7h8h As Long, rAM As Long, rSoir As Long, rNuit As Long, _
    rP0645 As Long, rP7H8H As Long, rP81630 As Long, rC15 As Long, rC20 As Long, rC20E As Long, rC19 As Long, _
    rNuitCode1 As Long, rNuitCode2 As Long, rNuitTotal As Long, rAction As Long, _
    cfgNuitCode1 As String, cfgNuitCode2 As String, _
    rLabelStart As Long, rLabelEnd As Long, rDates As Long, rDatesSrc As Long, colDebut As Long, colFin As Long, _
    labelFontName As String, labelFontSize As Long, labelFontBold As Long, _
    totalLabelFontSize As Long, totalLabelFontColor As Long)
    
    If lig1 > 0 Then ws.Cells(lig1, 1).value = "Matin"
    If lig2 > 0 Then ws.Cells(lig2, 1).value = "Apres-Midi"
    If lig3 > 0 Then ws.Cells(lig3, 1).value = "Soir"
    ws.Cells(rMeteo, 1).value = "METEO DU JOUR"
    ws.Cells(rMatin, 1).value = "Matin (Ecart)"
    ws.Cells(r7h8h, 1).value = "7h-8h (CRITIQUE)"
    ws.Cells(r7h8h, 1).Font.Bold = True
    ws.Cells(r7h8h, 1).Font.Color = vbRed
    ws.Cells(rAM, 1).value = "Apres-Midi (Ecart)"
    ws.Cells(rSoir, 1).value = "Soir (Ecart)"
    ws.Cells(rNuit, 1).value = "Nuit (Ecart)"
    If rP0645 > 0 Then ws.Cells(rP0645, 1).value = "Present a 06H45"
    If rP7H8H > 0 Then ws.Cells(rP7H8H, 1).value = "Presence entre 7h et 8h"
    If rP81630 > 0 Then ws.Cells(rP81630, 1).value = "Presence a 8 16h30"
    ws.Cells(rC15, 1).value = "Couverture C15"
    ws.Cells(rC20, 1).value = "Couverture C20"
    If rC20E > 0 Then ws.Cells(rC20E, 1).value = "Couverture C20 E"
    ws.Cells(rC19, 1).value = "Couverture C19"
    If rNuitCode1 > 0 Then ws.Cells(rNuitCode1, 1).value = cfgNuitCode1
    If rNuitCode2 > 0 Then ws.Cells(rNuitCode2, 1).value = cfgNuitCode2
    If rNuitTotal > 0 Then ws.Cells(rNuitTotal, 1).value = "Total Nuit"
    ws.Cells(rAction, 1).value = "DECISION"
    
    With ws.Range(ws.Cells(rLabelStart, 1), ws.Cells(rLabelEnd, 1))
        .Font.Name = labelFontName
        .Font.Size = labelFontSize
        .Font.Bold = (labelFontBold <> 0)
        .Interior.Color = RGB(242, 242, 242)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    If lig1 > 0 Then With ws.Cells(lig1, 1): .Font.Name = labelFontName: .Font.Size = totalLabelFontSize: .Font.Color = totalLabelFontColor: End With
    If lig2 > 0 Then With ws.Cells(lig2, 1): .Font.Name = labelFontName: .Font.Size = totalLabelFontSize: .Font.Color = totalLabelFontColor: End With
    If lig3 > 0 Then With ws.Cells(lig3, 1): .Font.Name = labelFontName: .Font.Size = totalLabelFontSize: .Font.Color = totalLabelFontColor: End With
End Sub


