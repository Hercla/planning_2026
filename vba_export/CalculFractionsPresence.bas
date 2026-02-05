Attribute VB_Name = "CalculFractionsPresence"
' ExportedAt: 2026-01-24 22:55:16 | Workbook: Planning_2026.xlsm
Option Explicit

' =========================================================================================
'   MACRO COMPLETE - TOUT EN UN
'   Pas de dependances externes
'   Date: 21 janvier 2026
' =========================================================================================

Sub Calculer_Totaux_Planning()
    Dim ws As Worksheet
    Dim wsCodesSpec As Worksheet
    Dim wsConfigCodes As Worksheet
    Dim wsConfig As Worksheet

    Set ws = ActiveSheet
    
    ' --- VALIDATION : Verifier qu'on est sur un onglet mois ---
    Dim nomOnglet As String
    nomOnglet = ws.Name
    If nomOnglet <> "Janv" And nomOnglet <> "Fev" And nomOnglet <> "Mars" And _
       nomOnglet <> "Avril" And nomOnglet <> "Mai" And nomOnglet <> "Juin" And _
       nomOnglet <> "Juil" And nomOnglet <> "Aout" And nomOnglet <> "Sept" And _
       nomOnglet <> "Oct" And nomOnglet <> "Nov" And nomOnglet <> "Dec" Then
        MsgBox "ERREUR : Cette macro ne peut être executee que sur un onglet mois." & vbCrLf & _
               "Selectionnez un onglet parmi : Janv, Fev, Mars, Avril, Mai, Juin, Juil, Aout, Sept, Oct, Nov, Dec", _
               vbExclamation, "Onglet invalide"
        Exit Sub
    End If
    ' -----------------------------------------------------------
    
    On Error Resume Next
    Set wsCodesSpec = ThisWorkbook.Sheets("Codes_Speciaux")
    Set wsConfigCodes = ThisWorkbook.Sheets("Config_Codes")
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    If wsCodesSpec Is Nothing And wsConfigCodes Is Nothing Then
        MsgBox "ERREUR : Feuilles 'Codes_Speciaux' ou 'Config_Codes' introuvables.", vbCritical
        Exit Sub
    End If

    ' --- CONFIGURATION ---
    Dim configGlobal As Object
    Set configGlobal = ChargerConfig(wsConfig)
    
    ' Parametres generaux - UTILISER LES CLES EXACTES DE Feuil_Config
    Dim ligneDebut As Long: ligneDebut = Module_Planning_Core.CfgLongFromDict(configGlobal, "CHK_FirstPersonnelRow", 6)
    Dim ligneFin As Long: ligneFin = Module_Planning_Core.CfgLongFromDict(configGlobal, "ligneFin", 28)
    Dim colDebut As Long: colDebut = Module_Planning_Core.CfgLongFromDict(configGlobal, "PLN_FirstDayCol", 3)
    Dim colFin As Long: colFin = Module_Planning_Core.CfgLongFromDict(configGlobal, "PLN_LastDayCol", 33)
    Dim ligneNumJour As Long: ligneNumJour = Module_Planning_Core.CfgLongFromDict(configGlobal, "PLN_Row_DayNumbers", 4)
    Dim couleurIgnore As Long: couleurIgnore = Module_Planning_Core.CfgLongFromDict(configGlobal, "CHK_IgnoreColor", 15849925)
    
    ' Couleurs d'exclusion speciales (clés corrigées)
    Dim couleurINFAdmin As Long: couleurINFAdmin = Module_Planning_Core.CfgLongFromDict(configGlobal, "COULEUR_INF_ADMIN", 65535)
    Dim couleurBleuClair As Long: couleurBleuClair = Module_Planning_Core.CfgLongFromDict(configGlobal, "COULEUR_BLEU_CLAIR", 15128749)
    
    ' Fonctions a compter (INF, AS, CEFA par defaut, ou depuis config)
    Dim fonctionsACompter As String: fonctionsACompter = GetCfgStr(configGlobal, "CHK_InfFunctions")
    If fonctionsACompter = "" Then fonctionsACompter = "INF,AS,CEFA"
    fonctionsACompter = Replace(fonctionsACompter, " ", "") ' Securite: supprimer les espaces
    fonctionsACompter = Replace(fonctionsACompter, ";", ",") ' Securite: remplacer ; par ,

    ' Lignes destination pour les totaux (lire depuis Feuil_Config)
    Dim lig(1 To 11) As Long
    lig(1) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_Matin", 60)    ' Matin
    lig(2) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_AM", 61)       ' Apres-midi (AM)
    lig(3) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_Soir", 62)     ' Soir
    lig(4) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_Nuit", 0)      ' Nuit (0 si pas defini)
    lig(5) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_P_0645", 64)   ' Present a 06H45
    lig(6) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_P_7H8H", 65)   ' Presence entre 7h et 8h
    lig(7) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_P_8H1630", 66) ' Presence a 16h30
    lig(8) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_C15", 67)      ' Presence en C 15
    lig(9) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_C20", 68)      ' Presence en C 20
    lig(10) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_C20E", 69)    ' Presence en C 20 E
    lig(11) = Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_C19", 70)     ' Presence en C 19
    ' --- COCKPIT LAYOUT (zone 60-73) ---
    Dim rDates As Long: rDates = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_DATES", 59)
    Dim rDatesSrc As Long: rDatesSrc = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_DATES_SRC", 3)
    Dim rLabelStart As Long: rLabelStart = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_LABEL_START", 59)
    Dim rLabelEnd As Long: rLabelEnd = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_LABEL_END", 75)
    Dim rMeteo As Long: rMeteo = 60
    Dim rMatin As Long: rMatin = 61
    Dim r7h8h As Long: r7h8h = 62
    Dim rAM As Long: rAM = 63
    Dim rSoir As Long: rSoir = 64
    Dim rNuit As Long: rNuit = 65
    Dim rSep As Long: rSep = 66
    Dim rP0645 As Long: rP0645 = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_P0645", 64)
    Dim rP7H8H As Long: rP7H8H = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_P7H8H", 65)
    Dim rP81630 As Long: rP81630 = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_P81630", 66)
    Dim rC15 As Long: rC15 = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_C15", Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_C15", 67))
    Dim rC20 As Long: rC20 = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_C20", Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_C20", 68))
    Dim rC20E As Long: rC20E = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_C20E", Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_C20E", 69))
    Dim rC19 As Long: rC19 = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_C19", Module_Planning_Core.CfgLongFromDict(configGlobal, "CALC_ROW_C19", 70))
    Dim rNuitCode1 As Long: rNuitCode1 = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_NUIT_CODE1", 71)
    Dim rNuitCode2 As Long: rNuitCode2 = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_NUIT_CODE2", 72)
    Dim rNuitTotal As Long: rNuitTotal = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_NUIT_TOTAL", 73)
    Dim rAction As Long: rAction = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_ROW_ACTION", 75)

    Dim target78 As Long: target78 = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_7H8H", 2)

    Dim cfgSpecialC20 As String: cfgSpecialC20 = GetCfgStr(configGlobal, "COCKPIT_SPECIAL_C20_DAYS")
    If cfgSpecialC20 = "" Then cfgSpecialC20 = "VEN,SAM,FERIE"
    Dim cfgC15Required As String: cfgC15Required = GetCfgStr(configGlobal, "COCKPIT_C15_REQUIRED_DAYS")
    If cfgC15Required = "" Then cfgC15Required = "LUN,MAR,MER,JEU,DIM"
    Dim cfgC15Forbidden As String: cfgC15Forbidden = GetCfgStr(configGlobal, "COCKPIT_C15_FORBIDDEN_DAYS")
    If cfgC15Forbidden = "" Then cfgC15Forbidden = "VEN,SAM,FERIE"
    Dim cfgC20ReqSpecial As Long: cfgC20ReqSpecial = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_C20_REQUIRED_SPECIAL", 2)
    Dim cfgC20ReqNormal As Long: cfgC20ReqNormal = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_C20_REQUIRED_NORMAL", 1)
    Dim cfgC19Req As Long: cfgC19Req = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_C19_REQUIRED", 1)
    Dim cfgC15ReqCount As Long: cfgC15ReqCount = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_C15_REQUIRED_COUNT", 1)
    Dim cfgCoupeMinTotal As Long: cfgCoupeMinTotal = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_COUPES_MIN_TOTAL", 3)
    Dim cfg78MinTotal As Long: cfg78MinTotal = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_7H8H_MIN_TOTAL", target78)
    Dim cfg78MinInf As Long: cfg78MinInf = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_7H8H_MIN_INF", 2)
    Dim cfg78MinAS As Long: cfg78MinAS = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_7H8H_MIN_AS", 1)
    Dim cfg78AllowAllInf As Long: cfg78AllowAllInf = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_7H8H_ALLOW_ALL_INF", 1)
    Dim cfgP0645Req As Long: cfgP0645Req = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_P0645_REQUIRED", 1)
    Dim cfgCheckFontOk As Long: cfgCheckFontOk = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_CHECK_FONT_SIZE_OK", 14)
    Dim cfgCheckFontAlert As Long: cfgCheckFontAlert = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_CHECK_FONT_SIZE_ALERT", 9)
    Dim cfgCheckFontName As String: cfgCheckFontName = GetCfgStr(configGlobal, "COCKPIT_CHECK_FONT_NAME")
    If cfgCheckFontName = "" Then cfgCheckFontName = "Arial Narrow"
    Dim cfgNuitCode1 As String: cfgNuitCode1 = GetCfgStr(configGlobal, "COCKPIT_NUIT_CODE_1")
    If cfgNuitCode1 = "" Then cfgNuitCode1 = "19:45 6:45"
    Dim cfgNuitCode2 As String: cfgNuitCode2 = GetCfgStr(configGlobal, "COCKPIT_NUIT_CODE_2")
    If cfgNuitCode2 = "" Then cfgNuitCode2 = "20 7"
    Dim cfgNuitReqTotal As Long: cfgNuitReqTotal = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_NUIT_REQUIRED_TOTAL", 2)
    Dim cfgLabelFontName As String: cfgLabelFontName = GetCfgStr(configGlobal, "COCKPIT_LABEL_FONT_NAME")
    If cfgLabelFontName = "" Then cfgLabelFontName = "Arial Narrow"
    Dim cfgLabelFontSize As Long: cfgLabelFontSize = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_LABEL_FONT_SIZE", 16)
    Dim cfgLabelFontBold As Long: cfgLabelFontBold = Module_Planning_Core.CfgLongFromDict(configGlobal, "COCKPIT_LABEL_FONT_BOLD", 0)

    Dim rClearStart As Long: rClearStart = rMeteo
    If rDates < rClearStart Then rClearStart = rDates
    If rLabelStart < rClearStart Then rClearStart = rLabelStart
    Dim rClearEnd As Long: rClearEnd = rAction
    If rP0645 > rClearEnd Then rClearEnd = rP0645
    If rP7H8H > rClearEnd Then rClearEnd = rP7H8H
    If rP81630 > rClearEnd Then rClearEnd = rP81630
    If rC20E > rClearEnd Then rClearEnd = rC20E
    If rNuitCode1 > rClearEnd Then rClearEnd = rNuitCode1
    If rNuitCode2 > rClearEnd Then rClearEnd = rNuitCode2
    If rNuitTotal > rClearEnd Then rClearEnd = rNuitTotal
    With ws.Range(ws.Cells(rClearStart, colDebut), ws.Cells(rClearEnd + 1, colFin))
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
        .Borders.LineStyle = xlNone
    End With

    If rDatesSrc > 0 Then
        ws.Range(ws.Cells(rDates, colDebut), ws.Cells(rDates, colFin)).value = _
            ws.Range(ws.Cells(rDatesSrc, colDebut), ws.Cells(rDatesSrc, colFin)).value
    End If

    

    
    ' Effectifs prevus par type de jour
    Dim effSem(1 To 4) As Long, effWE(1 To 4) As Long, effFER(1 To 4) As Long
    effSem(1) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_SEM_Matin", 7)
    effSem(2) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_SEM_PM", 3)
    effSem(3) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_SEM_Soir", 3)
    effSem(4) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_SEM_Nuit", 2)
    effWE(1) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_WE_Matin", 5)
    effWE(2) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_WE_PM", 2)
    effWE(3) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_WE_Soir", 3)
    effWE(4) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_WE_Nuit", 2)
    effFER(1) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_FER_Matin", 5)
    effFER(2) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_FER_PM", 2)
    effFER(3) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_FER_Soir", 3)
    effFER(4) = Module_Planning_Core.CfgLongFromDict(configGlobal, "EFF_FER_Nuit", 2)
    
    ' Seuil minimum INF
    Dim seuilMinINF As Long: seuilMinINF = Module_Planning_Core.CfgLongFromDict(configGlobal, "ALERT_SEUIL_MIN_INF", 2)
    
    ' Annee pour calcul feries
    Dim Annee As Long: Annee = Module_Planning_Core.CfgLongFromDict(configGlobal, "CFG_Year", Year(Date))
    Dim joursFeries As Object
    Set joursFeries = BuildFeriesBE(Annee)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- CHARGER CODES ---
    Dim dictCodes As Object
    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare

    If Not wsCodesSpec Is Nothing Then ChargerSpeciaux wsCodesSpec, dictCodes
    If Not wsConfigCodes Is Nothing Then ChargerConfigCodes wsConfigCodes, dictCodes, configGlobal
    
    ' --- CHARGER FONCTIONS DU PERSONNEL ---
    Dim dictFonctions As Object
    Set dictFonctions = ChargerFonctionsPersonnel()
    
    ' --- CHARGER CEFA EN FORMATION (% jaune dans Personnel) ---
    Dim dictCEFAFormation As Object
    Set dictCEFAFormation = ChargerCEFAEnFormation(ws.Name, couleurINFAdmin)
    
    ' --- CHARGER EXCEPTIONS D'EXCLUSION (Config_Exceptions = source unique) ---
    Dim arrExclusions As Variant
    Dim nbExclusions As Long
    arrExclusions = ChargerExclusionsCalcul()
    If IsArray(arrExclusions) Then nbExclusions = UBound(arrExclusions, 1) Else nbExclusions = 0
    


    ' --- CALCULS ---
    Dim col As Long, i As Long, j As Long
    Dim cell As Range, code As String, nomPersonne As String
    Dim tot(1 To 11) As Double
    Dim totINF(1 To 11) As Double
    Dim totAS(1 To 11) As Double
    Dim vals As Variant
    Dim estINF As Boolean
    Dim estAS As Boolean
    Dim c19INF As Long, c19AS As Long, c20Cnt As Long, c20ECnt As Long, c15Cnt As Long
    Dim nuitCode1Cnt As Long, nuitCode2Cnt As Long
    
    ' Variables pour alertes
    Dim colorAlerte As Long: colorAlerte = RGB(255, 100, 100)
    Dim colorWarning As Long: colorWarning = RGB(255, 200, 100)
    Dim colorINFManque As Long: colorINFManque = RGB(255, 0, 0)
    Dim numJour As Long, dateJour As Date, wd As Integer
    Dim estWE As Boolean, estFerie As Boolean
    Dim effPrevu(1 To 4) As Long
    Dim nomMois As String: nomMois = ws.Name
    Dim k As Integer, val1630 As Double, valC20E As Double
    
    ' Compteurs pour resume
    Dim joursAlerteINF As Long: joursAlerteINF = 0
    Dim joursManqueEff As Long: joursManqueEff = 0
    Dim joursProbleme1630 As Long: joursProbleme1630 = 0
    Dim alerteINFjour As Boolean, manqueEffJour As Boolean


    For col = colDebut To colFin
        For j = 1 To 11: tot(j) = 0: totINF(j) = 0: totAS(j) = 0: Next j
        c19INF = 0: c19AS = 0: c20Cnt = 0: c20ECnt = 0: c15Cnt = 0
        nuitCode1Cnt = 0: nuitCode2Cnt = 0

        For i = ligneDebut To ligneFin
            Set cell = ws.Cells(i, col)
            If cell.Interior.Color <> couleurIgnore Then
                code = Trim(CStr(cell.value))
                If code <> "" Then
                    ' Recuperer la fonction de la personne (pour comptage INF separe)
                    nomPersonne = Trim(CStr(ws.Cells(i, 1).value))
                    Dim fonctionPersonne As String: fonctionPersonne = ""
                    If dictFonctions.Exists(Replace(nomPersonne, " ", "_")) Then
                        fonctionPersonne = UCase(dictFonctions(Replace(nomPersonne, " ", "_")))
                    ElseIf dictFonctions.Exists(nomPersonne) Then
                        fonctionPersonne = UCase(dictFonctions(nomPersonne))
                    End If
                    
                    ' Couleur de la cellule pour exclusions
                    Dim couleurCellule As Long: couleurCellule = cell.Interior.Color
                    
                    ' PAR DEFAUT: Utiliser la liste definie dans Feuil_Config (CHK_InfFunctions)
                    ' ex: "INF,AS,CEFA". Si absent de la liste -> False.
                    Dim compterDansTotaux As Boolean: compterDansTotaux = False
                    
                    Dim fctUpper As String: fctUpper = UCase(fonctionPersonne)
                    
                    ' Verification si la fonction est dans la liste configuree (encadre de virgules pour match exact)
                    Dim estFonctionAutorisee As Boolean
                    estFonctionAutorisee = (InStr("," & UCase(fonctionsACompter) & ",", "," & fctUpper & ",") > 0)
                    
                    ' PAR DEFAUT: Ne pas compter (sauf si fonction autorisee)
                    compterDansTotaux = False

                    
                    If estFonctionAutorisee Then
                        If fctUpper = "CEFA" Then
                            ' CEFA : On compte SAUF si en formation (jaune dans Personnel)
                            Dim cleCefa As String
                            If dictFonctions.Exists(Replace(nomPersonne, " ", "_")) Then
                                cleCefa = Replace(nomPersonne, " ", "_")
                            Else
                                cleCefa = nomPersonne
                            End If
                            
                            If dictCEFAFormation.Exists(cleCefa) Then
                                compterDansTotaux = False ' En formation ce mois-ci -> Exclu
                            Else
                                compterDansTotaux = True ' Pas en formation -> Inclus
                            End If
                        Else
                            ' Pour INF, AS, etc. (si dans la liste) -> True
                            compterDansTotaux = True
                        End If
                    End If
                    

                    
                    ' Exclure si couleur d'absence (codes comme WE, CA, MAL, etc.)
                    Dim codeUpper As String: codeUpper = UCase(code)
                    If codeUpper = "WE" Or codeUpper Like "MAL*" Or codeUpper Like "CA*" _
                       Or codeUpper Like "RCT*" Or codeUpper Like "MAT*" Or codeUpper Like "MUT*" _
                       Or codeUpper = "CTR" Or codeUpper = "DP" Or codeUpper = "RHS" _
                       Or codeUpper = "EL" Or codeUpper Like "AFC*" _
                       Or codeUpper Like "3/4*" Or codeUpper Like "4/5*" _
                       Or codeUpper Like "CP*" Then
                        compterDansTotaux = False
                    End If
                    
                    ' Exclure si couleur ignore (config)
                    If couleurCellule = couleurIgnore Then compterDansTotaux = False
                    
                    ' Exclure si couleur jaune (administration/formation INF ou CEFA)
                    ' Ces personnes ne doivent PAS etre comptees dans les totaux
                    If couleurCellule = couleurINFAdmin Then compterDansTotaux = False
                    
                    ' Exclure si couleur bleu clair (bains les mercredis - Mamadou, Edelyne)
                    ' Ces personnes font les bains et ne comptent pas dans les effectifs soignants
                    If couleurCellule = couleurBleuClair Then compterDansTotaux = False
                    
                    ' Determiner le jour de la semaine (utilise pour les exclusions)
                    Dim jourActuel As String: jourActuel = ""
                    On Error Resume Next
                    jourActuel = GetJourNom(ws.Cells(ligneNumJour, col).value, Annee, nomMois)
                    On Error GoTo 0
                    
                    ' === VERIFICATION EXCLUSIONS CONFIG_EXCEPTIONS (source unique) ===
                    If nbExclusions > 0 And compterDansTotaux Then
                        Dim exIdx As Long
                        Dim exNom As String, exCode As String, exJours As String, exCouleur As String
                        
                        For exIdx = 1 To nbExclusions
                            exNom = CStr(arrExclusions(exIdx, 1))
                            exCode = CStr(arrExclusions(exIdx, 2))
                            exJours = CStr(arrExclusions(exIdx, 3))
                            exCouleur = UCase(CStr(arrExclusions(exIdx, 6)))
                            
                            ' Verifier si la regle match (avec ou sans couleur)
                            Dim nomMatch As Boolean, codeMatch As Boolean, jourMatch As Boolean, couleurMatch As Boolean
                            nomMatch = (exNom = "" Or exNom = "*" Or nomPersonne Like exNom)
                            codeMatch = (exCode = "" Or code Like exCode Or code = exCode)
                            jourMatch = (exJours = "" Or InStr(exJours, Left(jourActuel, 3)) > 0)
                            
                            ' Si couleur specifiee, verifier; sinon match automatique
                            If exCouleur <> "" Then
                                couleurMatch = MatchCouleur(couleurCellule, exCouleur)
                            Else
                                couleurMatch = True
                            End If
                            
                            If nomMatch And codeMatch And jourMatch And couleurMatch Then
                                compterDansTotaux = False
                                Exit For
                            End If
                        Next exIdx
                    End If
                    
                    ' === OBTENIR LES VALEURS DE FRACTIONS ===
                    Dim codeFound As Boolean: codeFound = False
                    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
                    Dim vLocal(1 To 11) As Double
                    Dim jj As Long
                    For jj = 1 To 11: vLocal(jj) = 0: Next jj
                    
                    ' 1) Essayer dictionnaire exact
                    If dictCodes.Exists(code) Then
                        vals = dictCodes(code)
                        For jj = 1 To 11: vLocal(jj) = vals(jj): Next jj
                        codeFound = True
                    ' 2) Sinon, parser le code comme horaire
                    ElseIf ParseCode(code, h1, f1, h2, f2) Then
                        CalcPeriodes h1, f1, h2, f2, vLocal(1), vLocal(2), vLocal(3), vLocal(4)
                        CalcPresSpec h1, f1, h2, f2, vLocal(5), vLocal(6), vLocal(7)
                        ' Detection C15, C20, C20E, C19 par horaires
                        If IsCodeC15(h1, f1, h2, f2) Then vLocal(8) = 1
                        If IsCodeC20(h1, f1, h2, f2) Then vLocal(9) = 1
                        If IsCodeC20E(h1, f1, h2, f2) Then vLocal(10) = 1
                        If IsCodeC19(h1, f1, h2, f2) Then vLocal(11) = 1
                        codeFound = True
                    End If
                    
                    If codeFound Then
                        ' Compter dans les totaux seulement si eligible
                        If compterDansTotaux Then
                            For jj = 1 To 11: tot(jj) = tot(jj) + vLocal(jj): Next jj
                            If vLocal(8) > 0 Then c15Cnt = c15Cnt + 1
                            If vLocal(9) > 0 Then c20Cnt = c20Cnt + 1
                            If vLocal(10) > 0 Then c20ECnt = c20ECnt + 1
                            If vLocal(11) > 0 Then
                                If fctUpper = "INF" Then
                                    c19INF = c19INF + 1
                                ElseIf fctUpper = "AS" Or fctUpper = "CEFA" Then
                                    c19AS = c19AS + 1
                                End If
                            End If
                            Dim codeU As String: codeU = UCase(Trim(code))
                            If codeU = UCase(cfgNuitCode1) Then nuitCode1Cnt = nuitCode1Cnt + 1
                            If codeU = UCase(cfgNuitCode2) Then nuitCode2Cnt = nuitCode2Cnt + 1
                        End If
                        
                        ' Compter separement les INF/AS (sauf si jaune = admin pour INF)
                        estINF = (UCase(fonctionPersonne) = "INF") And (couleurCellule <> couleurINFAdmin)
                        estAS = (fctUpper = "AS")
                        If estINF And compterDansTotaux Then
                            For jj = 1 To 11: totINF(jj) = totINF(jj) + vLocal(jj): Next jj
                        End If
                        If estAS And compterDansTotaux Then
                            For jj = 1 To 11: totAS(jj) = totAS(jj) + vLocal(jj): Next jj
                        End If
                    End If
                End If
            End If
            
        Next i

        ' --- COCKPIT RENDER ---
        numJour = 0
        On Error Resume Next
        numJour = CLng(ws.Cells(ligneNumJour, col).value)
        On Error GoTo 0

        If numJour > 0 And numJour <= 31 Then
            dateJour = DateFromMoisNom(nomMois, numJour, Annee)

            wd = Weekday(dateJour, vbMonday)
            estWE = (wd >= 6)
            estFerie = EstDansFeries(dateJour, joursFeries)
        Else
            estWE = False
            estFerie = False
            wd = 0
        End If

        If estFerie Then
            effPrevu(1) = effFER(1): effPrevu(2) = effFER(2)
            effPrevu(3) = effFER(3): effPrevu(4) = effFER(4)
        ElseIf estWE Then
            effPrevu(1) = effWE(1): effPrevu(2) = effWE(2)
            effPrevu(3) = effWE(3): effPrevu(4) = effWE(4)
        Else
            effPrevu(1) = effSem(1): effPrevu(2) = effSem(2)
            effPrevu(3) = effSem(3): effPrevu(4) = effSem(4)
        End If

        Dim jourRouge As Boolean: jourRouge = False
        Dim delta As Double

        delta = tot(1) - effPrevu(1)
        DrawDelta ws.Cells(rMatin, col), delta, False
        If delta >= 0 Then
            With ws.Cells(rMatin, col)
                .value = tot(1)
                .Font.Color = RGB(160, 160, 160)
                .Font.Bold = False
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
            End With
        End If
        If delta < 0 Then jourRouge = True

        delta = tot(2) - effPrevu(2)
        DrawDelta ws.Cells(rAM, col), delta, False
        If delta >= 0 Then
            With ws.Cells(rAM, col)
                .value = tot(2)
                .Font.Color = RGB(160, 160, 160)
                .Font.Bold = False
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
            End With
        End If
        If delta < 0 Then jourRouge = True

        delta = tot(3) - effPrevu(3)
        DrawDelta ws.Cells(rSoir, col), delta, False
        If delta >= 0 Then
            With ws.Cells(rSoir, col)
                .value = tot(3)
                .Font.Color = RGB(160, 160, 160)
                .Font.Bold = False
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
            End With
        End If
        If delta < 0 Then jourRouge = True

        delta = tot(4) - effPrevu(4)
        DrawDelta ws.Cells(rNuit, col), delta, False
        If delta >= 0 Then
            With ws.Cells(rNuit, col)
                .value = tot(4)
                .Font.Color = RGB(160, 160, 160)
                .Font.Bold = False
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
            End With
        End If
        If delta < 0 Then jourRouge = True

        Dim total78 As Double: total78 = tot(6)
        Dim inf78 As Double: inf78 = totINF(6)
        Dim as78 As Double: as78 = totAS(6)
        Dim ok78 As Boolean
        If total78 < cfg78MinTotal Then
            ok78 = False
            delta = total78 - cfg78MinTotal
        ElseIf (cfg78AllowAllInf <> 0) And (inf78 >= cfg78MinTotal) Then
            ok78 = True
            delta = 0
        Else
            ok78 = (inf78 >= cfg78MinInf And as78 >= cfg78MinAS)
            If ok78 Then
                delta = 0
            Else
                delta = -1
            End If
        End If
        DrawDelta ws.Cells(r7h8h, col), delta, True
        If Not ok78 Then jourRouge = True

        If rP0645 > 0 Then
            If tot(5) > 0 Then
                ws.Cells(rP0645, col).value = tot(5)
                ws.Cells(rP0645, col).Font.Size = cfgCheckFontOk
                ws.Cells(rP0645, col).Font.Name = cfgCheckFontName
                ws.Cells(rP0645, col).HorizontalAlignment = xlCenter
            Else
                ws.Cells(rP0645, col).value = ""
            End If
            If tot(5) < cfgP0645Req Then
                ws.Cells(rP0645, col).Interior.Color = RGB(255, 200, 150)
                ws.Cells(rP0645, col).Font.Size = cfgCheckFontAlert
                ws.Cells(rP0645, col).Font.Name = cfgCheckFontName
                ws.Cells(rP0645, col).HorizontalAlignment = xlCenter
                If tot(5) = 0 Then ws.Cells(rP0645, col).value = "MANQUE"
                jourRouge = True
            End If
        End If
        If rP7H8H > 0 Then
            ws.Cells(rP7H8H, col).value = IIf(tot(6) > 0, tot(6), "")
            If tot(6) > 0 Then
                ws.Cells(rP7H8H, col).Font.Size = cfgCheckFontOk
                ws.Cells(rP7H8H, col).Font.Name = cfgCheckFontName
                ws.Cells(rP7H8H, col).HorizontalAlignment = xlCenter
            End If
        End If
        If rP81630 > 0 Then
            ws.Cells(rP81630, col).value = IIf(tot(7) > 0, tot(7), "")
            If tot(7) > 0 Then
                ws.Cells(rP81630, col).Font.Size = cfgCheckFontOk
                ws.Cells(rP81630, col).Font.Name = cfgCheckFontName
                ws.Cells(rP81630, col).HorizontalAlignment = xlCenter
            End If
        End If

        Dim c19Type As String
        If c19INF > 0 Then
            c19Type = "INF"
        ElseIf c19AS > 0 Then
            c19Type = "AS"
        Else
            c19Type = ""
        End If

        Dim dayTags As String: dayTags = GetDayTags(wd, estFerie)
        Dim isSpecialC20Day As Boolean: isSpecialC20Day = DayInList(dayTags, cfgSpecialC20)
        Dim isC15Forbidden As Boolean: isC15Forbidden = DayInList(dayTags, cfgC15Forbidden)
        Dim isC15Required As Boolean: isC15Required = DayInList(dayTags, cfgC15Required)

        Dim okC19 As Boolean: okC19 = (c19INF + c19AS) >= cfgC19Req
        Dim okC15 As Boolean: okC15 = True
        Dim okNoC15 As Boolean: okNoC15 = True
        Dim totalC20 As Long: totalC20 = c20Cnt + c20ECnt
        Dim totalCoupes As Long: totalCoupes = c15Cnt + c20Cnt + c20ECnt + c19INF + c19AS
        Dim isC19AS As Boolean: isC19AS = (c19Type = "AS")
        Dim isC19INF As Boolean: isC19INF = (c19Type = "INF")
        Dim hasInf81630 As Boolean: hasInf81630 = (totINF(7) > 0)
        Dim needC15 As Boolean: needC15 = False
        Dim needC20 As Boolean: needC20 = False
        Dim needC20E As Boolean: needC20E = False
        Dim needThirdCoupe As Boolean: needThirdCoupe = False

        If isSpecialC20Day Then
            If isC15Forbidden Then okNoC15 = (c15Cnt = 0)
            ' Sur jours speciaux: 2 coupés 20 requis, C15 interdit
            needC20E = isC19AS And Not hasInf81630   ' si C19 AS, au moins un C20E (sauf INF 8-16:30)
        Else
            ' Jours normaux:
            ' - si C19 AS/CEFA -> C15 + C20E
            ' - si C19 INF -> C20 + (3 coupés au total, C15 ou C20E en 3e)
            If isC15Required And isC19AS Then needC15 = True
            If isC19INF Then needC20 = True
            If isC19AS And Not hasInf81630 Then needC20E = True
            If isC19INF Then
                If totalCoupes < cfgCoupeMinTotal And (c20Cnt + c20ECnt) >= 1 Then
                    needThirdCoupe = True
                End If
            End If
        End If

        Dim c19Total As Long: c19Total = c19INF + c19AS
        If rC15 > 0 Then
            ws.Cells(rC15, col).value = IIf(c15Cnt > 0, c15Cnt, "")
            If c15Cnt > 0 Then
                ws.Cells(rC15, col).Font.Size = cfgCheckFontOk
                ws.Cells(rC15, col).Font.Name = cfgCheckFontName
                ws.Cells(rC15, col).HorizontalAlignment = xlCenter
            End If
        End If
        If rC20 > 0 Then
            ws.Cells(rC20, col).value = IIf(c20Cnt > 0, c20Cnt, "")
            If c20Cnt > 0 Then
                ws.Cells(rC20, col).Font.Size = cfgCheckFontOk
                ws.Cells(rC20, col).Font.Name = cfgCheckFontName
                ws.Cells(rC20, col).HorizontalAlignment = xlCenter
            End If
        End If
        If rC20E > 0 Then
            ws.Cells(rC20E, col).value = IIf(c20ECnt > 0, c20ECnt, "")
            If c20ECnt > 0 Then
                ws.Cells(rC20E, col).Font.Size = cfgCheckFontOk
                ws.Cells(rC20E, col).Font.Name = cfgCheckFontName
                ws.Cells(rC20E, col).HorizontalAlignment = xlCenter
            End If
        End If
        If rC19 > 0 Then
            ws.Cells(rC19, col).value = IIf(c19Total > 0, c19Total, "")
            If c19Total > 0 Then
                ws.Cells(rC19, col).Font.Size = cfgCheckFontOk
                ws.Cells(rC19, col).Font.Name = cfgCheckFontName
                ws.Cells(rC19, col).HorizontalAlignment = xlCenter
            End If
        End If

        If Not okC19 Then
                With ws.Cells(rC19, col)
                    .value = "MANQUE"
                    .Interior.Color = RGB(255, 200, 150)
                    .Font.Size = cfgCheckFontAlert
                    .Font.Name = cfgCheckFontName
                    .HorizontalAlignment = xlCenter
                End With
                jourRouge = True
            End If

        If isSpecialC20Day Then
            If totalC20 < cfgC20ReqSpecial Then
                With ws.Cells(rC20, col)
                    .value = "MANQUE C20"
                    .Interior.Color = RGB(255, 200, 150)
                    .Font.Size = cfgCheckFontAlert
                    .Font.Name = cfgCheckFontName
                    .HorizontalAlignment = xlCenter
                End With
                jourRouge = True
            End If
            If needC20E And c20ECnt < 1 Then
                With ws.Cells(rC20E, col)
                    .value = "MANQUE C20E"
                    .Interior.Color = RGB(255, 200, 150)
                    .Font.Size = cfgCheckFontAlert
                    .Font.Name = cfgCheckFontName
                    .HorizontalAlignment = xlCenter
                End With
                jourRouge = True
            End If
            If isC15Forbidden And Not okNoC15 Then
                With ws.Cells(rC15, col)
                    .value = "C15 INTERDIT"
                    .Interior.Color = RGB(255, 200, 150)
                    .Font.Size = cfgCheckFontAlert
                    .Font.Name = cfgCheckFontName
                    .HorizontalAlignment = xlCenter
                End With
                jourRouge = True
            End If
        Else
            If needC20 And c20Cnt < 1 Then
                With ws.Cells(rC20, col)
                    .value = "MANQUE C20"
                    .Interior.Color = RGB(255, 200, 150)
                    .Font.Size = cfgCheckFontAlert
                    .Font.Name = cfgCheckFontName
                    .HorizontalAlignment = xlCenter
                End With
                jourRouge = True
            End If
            If needC20E And c20ECnt < 1 Then
                With ws.Cells(rC20E, col)
                    .value = "MANQUE C20E"
                    .Interior.Color = RGB(255, 200, 150)
                    .Font.Size = cfgCheckFontAlert
                    .Font.Name = cfgCheckFontName
                    .HorizontalAlignment = xlCenter
                End With
                jourRouge = True
            End If
            If needC15 And c15Cnt < cfgC15ReqCount Then
                With ws.Cells(rC15, col)
                    .value = "MANQUE"
                    .Interior.Color = RGB(255, 200, 150)
                    .Font.Size = cfgCheckFontAlert
                    .Font.Name = cfgCheckFontName
                    .HorizontalAlignment = xlCenter
                End With
                jourRouge = True
            End If
            If needThirdCoupe And c15Cnt = 0 And c20ECnt = 0 Then
                With ws.Cells(rC15, col)
                    .value = "MANQUE"
                    .Interior.Color = RGB(255, 200, 150)
                    .Font.Size = cfgCheckFontAlert
                    .Font.Name = cfgCheckFontName
                    .HorizontalAlignment = xlCenter
                End With
                jourRouge = True
            End If
        End If

        Dim nuitTotal As Long: nuitTotal = nuitCode1Cnt + nuitCode2Cnt
        If rNuitCode1 > 0 Then
            ws.Cells(rNuitCode1, col).value = IIf(nuitCode1Cnt > 0, nuitCode1Cnt, "")
            If nuitCode1Cnt > 0 Then
                ws.Cells(rNuitCode1, col).Font.Size = cfgCheckFontOk
                ws.Cells(rNuitCode1, col).Font.Name = cfgCheckFontName
                ws.Cells(rNuitCode1, col).HorizontalAlignment = xlCenter
            End If
        End If
        If rNuitCode2 > 0 Then
            ws.Cells(rNuitCode2, col).value = IIf(nuitCode2Cnt > 0, nuitCode2Cnt, "")
            If nuitCode2Cnt > 0 Then
                ws.Cells(rNuitCode2, col).Font.Size = cfgCheckFontOk
                ws.Cells(rNuitCode2, col).Font.Name = cfgCheckFontName
                ws.Cells(rNuitCode2, col).HorizontalAlignment = xlCenter
            End If
        End If
        If rNuitTotal > 0 Then
            ws.Cells(rNuitTotal, col).value = IIf(nuitTotal > 0, nuitTotal, "")
            If nuitTotal > 0 Then
                ws.Cells(rNuitTotal, col).Font.Size = cfgCheckFontOk
                ws.Cells(rNuitTotal, col).Font.Name = cfgCheckFontName
                ws.Cells(rNuitTotal, col).HorizontalAlignment = xlCenter
            End If
        End If
        If cfgNuitReqTotal > 0 And nuitTotal < cfgNuitReqTotal Then
            With ws.Cells(rNuitTotal, col)
                .Interior.Color = RGB(255, 200, 150)
                .Font.Size = cfgCheckFontAlert
                .Font.Name = cfgCheckFontName
                .HorizontalAlignment = xlCenter
            End With
            jourRouge = True
        End If

        With ws.Cells(rMeteo, col)
            If jourRouge Then
                .value = "."
                .Font.Color = vbRed
            Else
                .value = ""
            End If
            .HorizontalAlignment = xlCenter
        End With

        If jourRouge Then
            With ws.Cells(rAction, col)
                .value = "APPEL"
                .Interior.Color = vbBlack
                .Font.Color = vbWhite
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .Font.Size = 8
            End With
        End If
    Next col

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' --- ETIQUETTES COCKPIT (colonne A) ---
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

    ApplyCockpitLabelStyle ws, rLabelStart, rLabelEnd, rDates, rDatesSrc, colDebut, colFin, cfgLabelFontName, cfgLabelFontSize, cfgLabelFontBold

    With ws.Range(ws.Cells(rMeteo, colDebut), ws.Cells(rClearEnd, colFin))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlDot
        .Borders(xlInsideVertical).Color = RGB(200, 200, 200)
    End With

    MsgBox "Calcul cockpit termine.", vbInformation
End Sub


' =========================================================================================
'   COCKPIT: RENDU DELTA
' =========================================================================================
Private Sub DrawDelta(rng As Range, delta As Double, isCritical As Boolean)
    If delta < 0 Then
        rng.value = delta
        rng.Font.Bold = True
        If isCritical Then
            rng.Interior.Color = RGB(255, 0, 0)
            rng.Font.Color = vbWhite
        Else
            rng.Interior.Color = RGB(255, 200, 200)
            rng.Font.Color = RGB(200, 0, 0)
        End If
    Else
        rng.value = ""
        rng.Font.Bold = False
        rng.Font.Color = vbBlack
        rng.Interior.ColorIndex = xlNone
    End If
    rng.HorizontalAlignment = xlCenter
End Sub

' --- COCKPIT: helpers jours ---
Private Function GetDayTags(wd As Integer, estFerie As Boolean) As String
    Dim tag As String
    Select Case wd
        Case 1: tag = "LUN"
        Case 2: tag = "MAR"
        Case 3: tag = "MER"
        Case 4: tag = "JEU"
        Case 5: tag = "VEN"
        Case 6: tag = "SAM"
        Case 7: tag = "DIM"
        Case Else: tag = ""
    End Select
    If estFerie Then
        If tag <> "" Then
            tag = tag & ",FERIE"
        Else
            tag = "FERIE"
        End If
    End If
    GetDayTags = tag
End Function

Private Function DayInList(tags As String, listStr As String) As Boolean
    DayInList = False
    If Trim(listStr) = "" Then Exit Function
    Dim list As String: list = UCase(listStr)
    list = Replace(list, ";", ",")
    list = Replace(list, " ", "")
    Dim tagsU As String: tagsU = "," & UCase(tags) & ","
    Dim arr As Variant, T As Variant, tok As String
    arr = Split(list, ",")
    For Each T In arr
        tok = Trim(CStr(T))
        If tok = "" Then GoTo NextTok
        If tok = "WE" Then
            If InStr(tagsU, ",SAM,") > 0 Or InStr(tagsU, ",DIM,") > 0 Then
                DayInList = True
                Exit Function
            End If
        ElseIf InStr(tagsU, "," & tok & ",") > 0 Then
            DayInList = True
            Exit Function
        End If
NextTok:
    Next T
End Function

' --- COCKPIT: mise en forme label + dates ---
Private Sub ApplyCockpitLabelStyle(ws As Worksheet, rLabelStart As Long, rLabelEnd As Long, rDates As Long, rDatesSrc As Long, colDebut As Long, colFin As Long, labelFontName As String, labelFontSize As Long, labelFontBold As Long)
    If rLabelEnd < rLabelStart Then Exit Sub
    With ws.Range(ws.Cells(rLabelStart, 1), ws.Cells(rLabelEnd, 1))
        .Font.Name = labelFontName
        .Font.Size = labelFontSize
        .Font.Bold = (labelFontBold <> 0)
        .Interior.Color = RGB(242, 242, 242)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(210, 210, 210)
        End With
    End With

    If rDates > 0 Then
        Dim rngDates As Range
        Set rngDates = ws.Range(ws.Cells(rDates, colDebut), ws.Cells(rDates, colFin))
        If rDatesSrc > 0 Then
            ws.Range(ws.Cells(rDatesSrc, colDebut), ws.Cells(rDatesSrc, colFin)).Copy
            rngDates.PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
        Else
            With rngDates
                .Font.Name = "Calibri"
                .Font.Size = 9
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.Color = RGB(230, 240, 250)
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Color = RGB(210, 210, 210)
                End With
            End With
        End If
    End If
End Sub

' =============================================================================
'   FONCTIONS UTILITAIRES
' =============================================================================

Private Function BuildFeriesBE(ByVal Annee As Long) As Object
    Dim feries As Object
    Set feries = CreateObject("Scripting.Dictionary")
    Dim paques As Date
    
    paques = CalculerPaques(Annee)
    
    On Error Resume Next
    feries.Add CStr(DateSerial(Annee, 1, 1)), True
    feries.Add CStr(paques + 1), True
    feries.Add CStr(DateSerial(Annee, 5, 1)), True
    feries.Add CStr(paques + 39), True
    feries.Add CStr(paques + 50), True
    feries.Add CStr(DateSerial(Annee, 7, 21)), True
    feries.Add CStr(DateSerial(Annee, 8, 15)), True
    feries.Add CStr(DateSerial(Annee, 11, 1)), True
    feries.Add CStr(DateSerial(Annee, 11, 11)), True
    feries.Add CStr(DateSerial(Annee, 12, 25)), True
    On Error GoTo 0
    
    Set BuildFeriesBE = feries
End Function

Private Function CalculerPaques(ByVal Annee As Long) As Date
    Dim a As Integer, b As Integer, c As Integer
    Dim d As Integer, e As Integer, f As Integer
    Dim g As Integer, h As Integer, i As Integer
    Dim k As Integer, l As Integer, m As Integer
    Dim Mois As Integer, jour As Integer
    
    a = Annee Mod 19
    b = Annee \ 100
    c = Annee Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    Mois = (h + l - 7 * m + 114) \ 31
    jour = ((h + l - 7 * m + 114) Mod 31) + 1
    
    CalculerPaques = DateSerial(Annee, Mois, jour)
End Function

Private Function EstDansFeries(ByVal d As Date, ByVal feries As Object) As Boolean
    EstDansFeries = feries.Exists(CStr(d))
End Function

Private Function DateFromMoisNom(ByVal nomMois As String, ByVal jour As Long, ByVal Annee As Long) As Date
    Dim moisNum As Integer
    Select Case LCase(Left(nomMois, 4))
        Case "janv": moisNum = 1
        Case "fev", "fevr": moisNum = 2
        Case "mars": moisNum = 3
        Case "avri": moisNum = 4
        Case "mai": moisNum = 5
        Case "juin": moisNum = 6
        Case "juil": moisNum = 7
        Case "aout", "aout": moisNum = 8
        Case "sept": moisNum = 9
        Case "oct", "octo": moisNum = 10
        Case "nov", "nove": moisNum = 11
        Case "dec", "dece": moisNum = 12
        Case Else: moisNum = 1
    End Select
    DateFromMoisNom = DateSerial(Annee, moisNum, jour)
End Function

Private Function ChargerFonctionsPersonnel() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    Dim wsPersonnel As Worksheet
    On Error Resume Next
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If wsPersonnel Is Nothing Then
        Set ChargerFonctionsPersonnel = d
        Exit Function
    End If
    
    Dim lr As Long, arr As Variant, i As Long
    Dim nom As String, prenom As String, cleNomPrenom As String, fonction As String
    
    lr = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
    If lr < 2 Then Set ChargerFonctionsPersonnel = d: Exit Function
    
    arr = wsPersonnel.Range("B2:E" & lr).value
    For i = 1 To UBound(arr, 1)
        nom = Trim(CStr(arr(i, 1)))
        prenom = Trim(CStr(arr(i, 2)))
        fonction = Trim(CStr(arr(i, 4)))
        
        cleNomPrenom = nom & "_" & prenom
        
        If cleNomPrenom <> "_" And Not d.Exists(cleNomPrenom) Then
            d.Add cleNomPrenom, fonction
        End If
    Next i
    
    Set ChargerFonctionsPersonnel = d
End Function

Private Function ChargerCEFAEnFormation(nomMois As String, couleurFormation As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    Dim wsPersonnel As Worksheet
    On Error Resume Next
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If wsPersonnel Is Nothing Then
        Set ChargerCEFAEnFormation = d
        Exit Function
    End If
    
    Dim colMoisPct As Long
    Dim moisNum As Long: moisNum = MoisNumero(nomMois)
    colMoisPct = 29 + (moisNum - 1) * 2
    
    Dim lr As Long, i As Long
    Dim nom As String, prenom As String, cleNomPrenom As String, fonction As String
    Dim cellPct As Range
    
    lr = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
    If lr < 2 Then Set ChargerCEFAEnFormation = d: Exit Function
    
    For i = 2 To lr
        fonction = UCase(Trim(CStr(wsPersonnel.Cells(i, "E").value)))
        
        If fonction = "CEFA" Then
            nom = Trim(CStr(wsPersonnel.Cells(i, "B").value))
            prenom = Trim(CStr(wsPersonnel.Cells(i, "C").value))
            cleNomPrenom = nom & "_" & prenom
            
            Set cellPct = wsPersonnel.Cells(i, colMoisPct)
            If cellPct.Interior.Color = couleurFormation Then
                d(cleNomPrenom) = True
            End If
        End If
    Next i
    
    Set ChargerCEFAEnFormation = d
End Function

Private Function MoisNumero(nomMois As String) As Long
    Select Case LCase(Left(nomMois, 4))
        Case "janv": MoisNumero = 1
        Case "fev", "fevr": MoisNumero = 2
        Case "mars": MoisNumero = 3
        Case "avri": MoisNumero = 4
        Case "mai": MoisNumero = 5
        Case "juin": MoisNumero = 6
        Case "juil": MoisNumero = 7
        Case "aout", "aout": MoisNumero = 8
        Case "sept": MoisNumero = 9
        Case "oct", "octo": MoisNumero = 10
        Case "nov", "nove": MoisNumero = 11
        Case "dec", "dece": MoisNumero = 12
        Case Else: MoisNumero = 1
    End Select
End Function

Private Function ChargerConfig(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If ws Is Nothing Then Set ChargerConfig = d: Exit Function
    
    Dim lr As Long, arr As Variant, i As Long
    Dim cle As String, valeur As Variant
    lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Set ChargerConfig = d: Exit Function
    
    arr = ws.Range("A2:B" & lr).value
    For i = 1 To UBound(arr, 1)
        cle = Trim(CStr(arr(i, 1)))
        valeur = arr(i, 2)
        
        If cle <> "" And Not d.Exists(cle) Then
            d(cle) = valeur
        End If
    Next i
    Set ChargerConfig = d
End Function

Private Function Module_Planning_Core.CfgLongFromDict(d As Object, k As String, def As Long) As Long
    If d.Exists(k) Then
        If IsNumeric(d(k)) Then GetCfgLong = CLng(d(k)): Exit Function
    End If
    GetCfgLong = def
End Function

Private Function GetCfgStr(d As Object, k As String) As String
    If d.Exists(k) Then GetCfgStr = CStr(d(k)) Else GetCfgStr = ""
End Function

Private Sub ChargerSpeciaux(ws As Worksheet, d As Object)
    Dim lr As Long, arr As Variant, i As Long, k As Long
    Dim code As String, cc As String
    Dim v(1 To 11) As Double
    
    lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Exit Sub
    arr = ws.Range("A2:G" & lr).value
    
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" And Not d.Exists(code) Then
            For k = 1 To 11: v(k) = 0: Next k
            
            v(5) = NumVal(arr(i, 2))
            v(6) = NumVal(arr(i, 3))
            v(1) = NumVal(arr(i, 4))
            v(2) = NumVal(arr(i, 5))
            v(3) = NumVal(arr(i, 6))
            v(4) = NumVal(arr(i, 7))
            
            cc = UCase(Replace(code, " ", ""))
            If cc Like "C19*" Then v(11) = 1: If v(6) = 0 Then v(6) = 1
            If cc Like "C20E*" Then v(10) = 1
            If cc Like "C20*" And Not cc Like "C20E*" Then v(9) = 1
            If cc Like "C15*" Then v(8) = 1
            
            d.Add code, v
        End If
    Next i
End Sub

Private Sub ChargerConfigCodes(ws As Worksheet, d As Object, cfg As Object)
    Dim lr As Long, arr As Variant, i As Long, k As Long
    Dim code As String
    Dim v(1 To 11) As Double
    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
    Dim hasExtendedCols As Boolean
    Dim hStart As String, hPauseS As String, hPauseE As String, hEnd As String
    Dim manF645 As Variant, manF78 As Variant
    Dim manMatin As Variant, manPM As Variant, manSoir As Variant, manNuit As Variant
    
    Dim sC15 As String: sC15 = GetCfgStr(cfg, "SPECIAL_C15")
    Dim sC20 As String: sC20 = GetCfgStr(cfg, "SPECIAL_C20")
    Dim sC20E As String: sC20E = GetCfgStr(cfg, "SPECIAL_C20E")
    Dim sC19 As String: sC19 = GetCfgStr(cfg, "SPECIAL_C19")
    
    lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then Exit Sub
    
    hasExtendedCols = (ws.Cells(1, 6).value <> "")
    
    If hasExtendedCols Then
        arr = ws.Range("A2:O" & lr).value
    Else
        arr = ws.Range("A2:A" & lr).value
    End If
    
    For i = 1 To UBound(arr, 1)
        code = Trim(CStr(arr(i, 1)))
        If code <> "" And Not d.Exists(code) Then
            For k = 1 To 11: v(k) = 0: Next k
            h1 = 0: f1 = 0: h2 = 0: f2 = 0
            
            If hasExtendedCols Then
                manF645 = arr(i, 10)
                manF78 = arr(i, 11)
                manMatin = arr(i, 12)
                manPM = arr(i, 13)
                manSoir = arr(i, 14)
                manNuit = arr(i, 15)
                
                If IsNumeric(manMatin) And manMatin <> "" Then v(1) = CDbl(manMatin)
                If IsNumeric(manPM) And manPM <> "" Then v(2) = CDbl(manPM)
                If IsNumeric(manSoir) And manSoir <> "" Then v(3) = CDbl(manSoir)
                If IsNumeric(manNuit) And manNuit <> "" Then v(4) = CDbl(manNuit)
                
                If IsNumeric(manF645) And manF645 <> "" Then v(5) = CDbl(manF645)
                If IsNumeric(manF78) And manF78 <> "" Then v(6) = CDbl(manF78)
                
                hStart = Trim(CStr(arr(i, 6)))
                hPauseS = Trim(CStr(arr(i, 7)))
                hPauseE = Trim(CStr(arr(i, 8)))
                hEnd = Trim(CStr(arr(i, 9)))
                
                If hStart <> "" And hEnd <> "" Then
                    h1 = HeureDec(hStart)
                    If hPauseS <> "" And hPauseE <> "" Then
                        f1 = HeureDec(hPauseS)
                        h2 = HeureDec(hPauseE)
                        f2 = HeureDec(hEnd)
                    Else
                        f1 = HeureDec(hEnd)
                    End If
                    
                    Dim autoMat As Double, autoPM As Double
                    Dim autoSoir As Double, autoNuit As Double
                    CalcPeriodes h1, f1, h2, f2, autoMat, autoPM, autoSoir, autoNuit
                    
                    ' FIX: Les codes "Coupe" (C ...) ne doivent pas compter en PM par defaut.
                    ' Si manPM est vide, on veut 0, pas le calcul auto qui pourrait trouver 1.
                    Dim uc As String: uc = UCase(code)
                    If Left(uc, 1) = "C" Or Left(uc, 4) = "SA C" Or Left(uc, 4) = "DI C" Then
                        autoPM = 0
                    End If
                    
                    ' FIX USER: C 19 doit compter en Matin et Soir (et detecte en C 19)
                    If uc = "C 19" Or uc = "C 19 SA" Or uc = "C 19 DI" Then
                        If autoMat = 0 Then autoMat = 1
                        If autoSoir = 0 Then autoSoir = 1 ' Force Soir meme si calcul auto modifie
                    End If
                    
                    ' Detection des codes speciaux (C 15, C 19, C 20...) pour les logs du bas (lignes 67-70 excel)
                    ' AMELIORE: Detection par nom OU par horaires
                    Dim cc As String: cc = Replace(uc, " ", "")
                    Dim isC20EName As Boolean, isC20Name As Boolean, isC15Name As Boolean
                    isC20EName = (cc Like "C20E*")
                    isC20Name = (cc Like "C20*") And Not isC20EName
                    isC15Name = (cc Like "C15*")
                    
                    ' Detection C19 par nom ou horaires (7h-11:30 + 15:30-19h)
                    If cc Like "C19*" Then v(11) = 1: If v(6) = 0 Then v(6) = 1
                    If IsCodeC19Pattern(h1, f1, h2, f2) Then v(11) = 1
                    
                    ' Detection C20E par nom ou horaires (finit entre 20h15 et 21h)
                    If isC20EName Then v(10) = 1
                    If IsCodeC20E(h1, f1, h2, f2) Then v(10) = 1
                    
                    ' Detection C20 par nom ou horaires (8-12 + 16-20)
                    If isC20Name Then v(9) = 1
                    If IsCodeC20Pattern(h1, f1, h2, f2) Then v(9) = 1
                    
                    ' Detection C15 par nom ou horaires (8-12:15 + 16:30-20:15 ou variantes)
                    If isC15Name Then v(8) = 1
                    If IsCodeC15Pattern(h1, f1, h2, f2) Then v(8) = 1
                    
                    If Not (IsNumeric(manMatin) And manMatin <> "") Then v(1) = autoMat
                    If Not (IsNumeric(manPM) And manPM <> "") Then v(2) = autoPM
                    If Not (IsNumeric(manSoir) And manSoir <> "") Then v(3) = autoSoir
                    If Not (IsNumeric(manNuit) And manNuit <> "") Then v(4) = autoNuit
                    
                    Dim autoF645 As Double, autoF78 As Double, autoP81630 As Double
                    CalcPresSpec h1, f1, h2, f2, autoF645, autoF78, autoP81630
                    
                    If Not (IsNumeric(manF645) And manF645 <> "") Then v(5) = autoF645
                    If Not (IsNumeric(manF78) And manF78 <> "") Then v(6) = autoF78
                    v(7) = autoP81630
                    
                    If MatchSpecial(h1, f1, h2, f2, sC15) Then v(8) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC20) Then v(9) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC20E) Then v(10) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC19) Then v(11) = 1

                    ' Priorite des codes pour eviter doubles comptages
                    If isC20EName Then
                        v(10) = 1: v(9) = 0: v(8) = 0
                    ElseIf isC20Name Then
                        v(9) = 1: v(10) = 0: v(8) = 0
                    ElseIf isC15Name Then
                        v(8) = 1: v(9) = 0: v(10) = 0
                    Else
                        If v(8) > 0 Then v(9) = 0: v(10) = 0
                        If v(9) > 0 Or v(10) > 0 Then v(8) = 0
                    End If
                End If
            End If
            
            If v(1) = 0 And v(2) = 0 And v(3) = 0 And v(4) = 0 Then
                If ParseCode(code, h1, f1, h2, f2) Then
                    CalcPeriodes h1, f1, h2, f2, v(1), v(2), v(3), v(4)
                    CalcPresSpec h1, f1, h2, f2, v(5), v(6), v(7)
                    cc = Replace(UCase(code), " ", "")
                    isC20EName = (cc Like "C20E*")
                    isC20Name = (cc Like "C20*") And Not isC20EName
                    isC15Name = (cc Like "C15*")
                    If MatchSpecial(h1, f1, h2, f2, sC15) Then v(8) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC20) Then v(9) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC20E) Then v(10) = 1
                    If MatchSpecial(h1, f1, h2, f2, sC19) Then v(11) = 1
                    ' Priorite des codes pour eviter doubles comptages
                    If isC20EName Then
                        v(10) = 1: v(9) = 0: v(8) = 0
                    ElseIf isC20Name Then
                        v(9) = 1: v(10) = 0: v(8) = 0
                    ElseIf isC15Name Then
                        v(8) = 1: v(9) = 0: v(10) = 0
                    Else
                        If v(8) > 0 Then v(9) = 0: v(10) = 0
                        If v(9) > 0 Or v(10) > 0 Then v(8) = 0
                    End If
                End If
            End If
            
            d.Add code, v
        End If
    Next i
End Sub

Private Function NumVal(x As Variant) As Double
    If IsNumeric(x) Then NumVal = CDbl(x) Else NumVal = 0
End Function

Private Function HeureDec(s As Variant) As Double
    Dim p() As String
    Dim strVal As String
    
    If IsEmpty(s) Or s = "" Then
        HeureDec = 0
        Exit Function
    End If
    
    If IsNumeric(s) Then
        If CDbl(s) < 1 And CDbl(s) > 0 Then
            HeureDec = CDbl(s) * 24
            Exit Function
        ElseIf CDbl(s) >= 1 And CDbl(s) < 25 Then
            HeureDec = CDbl(s)
            Exit Function
        End If
    End If
    
    strVal = CStr(s)
    If InStr(strVal, ":") > 0 Then
        p = Split(strVal, ":")
        On Error Resume Next
        HeureDec = CDbl(p(0)) + CDbl(p(1)) / 60
        On Error GoTo 0
    ElseIf IsNumeric(strVal) Then
        HeureDec = CDbl(strVal)
    Else
        HeureDec = 0
    End If
End Function

Private Function ParseCode(c As String, ByRef s1 As Double, ByRef e1 As Double, ByRef s2 As Double, ByRef e2 As Double) As Boolean
    s1 = 0: e1 = 0: s2 = 0: e2 = 0
    Dim p() As String, tmp As String
    tmp = Trim(Replace(Replace(c, vbLf, " "), vbCr, " "))
    Do While InStr(tmp, "  ") > 0: tmp = Replace(tmp, "  ", " "): Loop
    p = Split(tmp, " ")
    
    On Error GoTo Err1
    If UBound(p) = 1 Then
        s1 = HeureDec(p(0)): e1 = HeureDec(p(1))
        ParseCode = True
    ElseIf UBound(p) >= 3 Then
        s1 = HeureDec(p(0)): e1 = HeureDec(p(1))
        s2 = HeureDec(p(2)): e2 = HeureDec(p(3))
        ParseCode = True
    End If
    Exit Function
Err1:
    ParseCode = False
End Function

Private Sub CalcPeriodes(h1 As Double, f1 As Double, h2 As Double, f2 As Double, ByRef mat As Double, ByRef AM As Double, ByRef soi As Double, ByRef nui As Double)
    mat = 0: AM = 0: soi = 0: nui = 0
    
    ' Determiner l'heure de fin effective
    Dim fin As Double
    fin = IIf(f2 > 0, f2, f1)
    
    ' Si pas d'horaire valide, sortir
    If h1 = 0 And f1 = 0 Then Exit Sub
    
    ' =============================================================================
    ' REGLES DE CALCUL DES PERIODES (configurables dans Feuil_Config si besoin)
    ' =============================================================================
    
    ' MATIN : Presence si debut < 13h
    If h1 < 13 Then
        mat = 1
    ElseIf h2 > 0 And h2 < 13 Then
        ' Cas horaire coupe avec reprise le matin
        mat = 1
    End If
    
    ' APRES-MIDI (PM) : Presence si fin > 13h
    If fin > 13 Then
        AM = 1
    End If
    ' Si horaire coupe avec 2eme partie l'apres-midi
    If h2 > 0 And f2 > 13 Then
        AM = 1
    End If
    
    ' SOIR : Presence si fin > 16h30
    ' - Fin entre 16h30 et 17h30 = 0.5 (demi-soir)
    ' - Fin apres 17h30 = 1 (soir complet)
    If fin > 17.5 Then
        soi = 1
    ElseIf fin > 16.5 Then
        soi = 0.5
    End If
    
    ' NUIT : Si commence >= 19h30 ou finit tot le matin (<= 7h15)
    If h1 >= 19.5 Or (fin > 0 And fin <= 7.25) Then
        ' Si finit minuit (24h ou 0h), compte comme 0.5
        If Abs(fin - 24) < 0.1 Or fin = 0 Then
            nui = 0.5
        Else
            nui = 1
        End If
    End If
End Sub

Private Function Overlap(hd As Double, hf As Double, td As Double, tf As Double) As Double
    Dim os As Double, oe As Double
    os = Application.Max(hd, td): oe = Application.Min(hf, tf)
    If oe > os Then Overlap = oe - os Else Overlap = 0
End Function

Private Sub CalcPresSpec(h1 As Double, f1 As Double, h2 As Double, f2 As Double, ByRef p645 As Double, ByRef p78 As Double, ByRef p81630 As Double)
    p645 = 0: p78 = 0: p81630 = 0
    If h1 <= 6.75 Then p645 = 1
    If h1 < 8 And f1 > 7 Then p78 = 1
    If Abs(f1 - 16.5) < 0.25 Or Abs(f2 - 16.5) < 0.25 Then p81630 = 1
End Sub

Private Function MatchSpecial(h1 As Double, f1 As Double, h2 As Double, f2 As Double, def As String) As Boolean
    MatchSpecial = False
    If def = "" Then Exit Function
    Dim p() As String: p = Split(def, " ")
    If UBound(p) < 3 Then Exit Function
    Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double
    c1 = HeureDec(p(0)): c2 = HeureDec(p(1)): c3 = HeureDec(p(2)): c4 = HeureDec(p(3))
    Const T As Double = 0.02
    If Abs(h1 - c1) < T And Abs(f1 - c2) < T And Abs(h2 - c3) < T And Abs(f2 - c4) < T Then MatchSpecial = True
End Function

Private Function IsCodeC15(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC15 = False
    Dim finEffective As Double
    finEffective = IIf(f2 > 0, f2, f1)
    If finEffective >= 15 And finEffective <= 15.5 Then IsCodeC15 = True
End Function

Private Function IsCodeC20(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC20 = False
    Dim finEffective As Double
    finEffective = IIf(f2 > 0, f2, f1)
    If finEffective >= 19.75 And finEffective <= 20.25 Then IsCodeC20 = True
End Function

Private Function IsCodeC20E(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC20E = False
    Dim finEffective As Double
    finEffective = IIf(f2 > 0, f2, f1)
    If finEffective > 20.25 And finEffective <= 21 Then IsCodeC20E = True
End Function

Private Function IsCodeC19(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    IsCodeC19 = False
    Dim finEffective As Double
    finEffective = IIf(f2 > 0, f2, f1)
    If finEffective >= 18.75 And finEffective <= 19.25 Then IsCodeC19 = True
End Function

' =============================================================================
' NOUVELLES FONCTIONS: Detection par pattern horaire (pas juste par nom)
' Permet de reconnaitre "8:30 12:45 16:30 20:15" comme C15 meme sans prefixe "C"
' =============================================================================

Private Function IsCodeC15Pattern(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    ' C15 = horaire coupe avec fin vers 20h15 (+/- 30 min)
    ' Exemples: 8 12:15 16:30 20:15, 8:30 12:45 16:30 20:15
    IsCodeC15Pattern = False
    If h2 = 0 Or f2 = 0 Then Exit Function ' Doit etre un horaire coupe
    
    ' Fin entre 19h45 et 20h45
    If f2 >= 19.75 And f2 <= 20.75 Then
        ' Pause vers 12h-13h et reprise vers 16h-17h
        If f1 >= 11.5 And f1 <= 13 And h2 >= 15.5 And h2 <= 17 Then
            IsCodeC15Pattern = True
        End If
    End If
End Function

Private Function IsCodeC19Pattern(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    ' C19 = horaire coupe 7-11:30 + 15:30-19 (fin vers 19h)
    IsCodeC19Pattern = False
    If h2 = 0 Or f2 = 0 Then Exit Function ' Doit etre un horaire coupe
    
    ' Fin entre 18h45 et 19h15
    If f2 >= 18.75 And f2 <= 19.25 Then
        ' Debut tot (avant 8h) et pause vers 11h-12h
        If h1 <= 8 And f1 >= 11 And f1 <= 12 Then
            IsCodeC19Pattern = True
        End If
    End If
End Function

Private Function IsCodeC20Pattern(h1 As Double, f1 As Double, h2 As Double, f2 As Double) As Boolean
    ' C20 = horaire coupe 8-12 + 16-20 (fin vers 20h)
    IsCodeC20Pattern = False
    If h2 = 0 Or f2 = 0 Then Exit Function ' Doit etre un horaire coupe
    
    ' Fin entre 19h45 et 20h15 (pas C20E qui va plus loin)
    If f2 >= 19.75 And f2 <= 20.25 Then
        ' Pause vers 12h et reprise vers 16h
        If f1 >= 11.5 And f1 <= 12.5 And h2 >= 15.5 And h2 <= 16.5 Then
            IsCodeC20Pattern = True
        End If
    End If
End Function

Private Function ChargerExclusionsCalcul() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config_Exceptions")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ChargerExclusionsCalcul = Empty
        Exit Function
    End If
    
    Dim lr As Long
    lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lr < 2 Then
        ChargerExclusionsCalcul = Empty
        Exit Function
    End If
    
    ChargerExclusionsCalcul = ws.Range("A2:F" & lr).value
End Function

Private Function GetJourNom(numJour As Variant, Annee As Long, nomMois As String) As String
    On Error GoTo Err1
    Dim d As Date
    Dim moisNum As Long: moisNum = MoisNumero(nomMois)
    d = DateSerial(Annee, moisNum, CLng(numJour))
    
    Select Case Weekday(d, vbMonday)
        Case 1: GetJourNom = "LUN"
        Case 2: GetJourNom = "MAR"
        Case 3: GetJourNom = "MER"
        Case 4: GetJourNom = "JEU"
        Case 5: GetJourNom = "VEN"
        Case 6: GetJourNom = "SAM"
        Case 7: GetJourNom = "DIM"
    End Select
    Exit Function
Err1:
    GetJourNom = ""
End Function

Private Function MatchCouleur(couleurCellule As Long, nomCouleur As String) As Boolean
    MatchCouleur = False
    Dim couleurCible As Long
    
    Select Case UCase(nomCouleur)
        Case "BLEU": couleurCible = 16711680
        Case "BLEU_CLAIR": couleurCible = 16776960
        Case "ROUGE": couleurCible = 255
        Case "JAUNE": couleurCible = 65535
        Case "ORANGE": couleurCible = 49407
        Case "CYAN": couleurCible = 16776960
        Case "ROSE": couleurCible = 13408767
        Case "GRIS": couleurCible = 12632256
        Case Else: Exit Function
    End Select
    
    If Abs(couleurCellule - couleurCible) < 100000 Then MatchCouleur = True
    If couleurCellule = couleurCible Then MatchCouleur = True
End Function

' =============================================================================
' MACRO DE DEBUG - VERIFICATION DES FRACTIONS POUR UNE COLONNE
' =============================================================================
Sub Debug_Verifier_Colonne()
    Dim ws As Worksheet
    Dim wsCodesSpec As Worksheet
    Dim wsConfigCodes As Worksheet
    Dim wsConfig As Worksheet
    
    Set ws = ActiveSheet
    
    Dim colSelect As Long
    colSelect = ActiveCell.Column
    
    If colSelect < 3 Or colSelect > 33 Then
        MsgBox "Placez le curseur sur une colonne de jour (C a AG)", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsCodesSpec = ThisWorkbook.Sheets("Codes_Speciaux")
    Set wsConfigCodes = ThisWorkbook.Sheets("Config_Codes")
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    Dim configGlobal As Object
    If Not wsConfig Is Nothing Then
        Set configGlobal = ChargerConfig(wsConfig)
    Else
        Set configGlobal = CreateObject("Scripting.Dictionary")
    End If
    
    Dim ligneDebut As Long: ligneDebut = Module_Planning_Core.CfgLongFromDict(configGlobal, "CHK_FirstPersonnelRow", 6)
    Dim ligneFin As Long: ligneFin = Module_Planning_Core.CfgLongFromDict(configGlobal, "ligneFin", 28)
    Dim ligneNumJour As Long: ligneNumJour = Module_Planning_Core.CfgLongFromDict(configGlobal, "PLN_Row_DayNumbers", 4)
    Dim couleurIgnore As Long: couleurIgnore = Module_Planning_Core.CfgLongFromDict(configGlobal, "CHK_IgnoreColor", 15849925)
    
    Dim dictCodes As Object
    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    
    If Not wsCodesSpec Is Nothing Then ChargerSpeciaux wsCodesSpec, dictCodes
    If Not wsConfigCodes Is Nothing Then ChargerConfigCodes wsConfigCodes, dictCodes, configGlobal
    
    Dim dictFonctions As Object
    Set dictFonctions = ChargerFonctionsPersonnel()
    
    ' FONCTIONS A COMPTER
    Dim fonctionsACompter As String: fonctionsACompter = GetCfgStr(configGlobal, "CHK_InfFunctions")
    If fonctionsACompter = "" Then fonctionsACompter = "INF,AS,CEFA"
    fonctionsACompter = Replace(fonctionsACompter, " ", "")
    fonctionsACompter = Replace(fonctionsACompter, ";", ",")
    
    ' CHARGER EXCEPTIONS
    Dim arrExclusions As Variant
    Dim nbExclusions As Long
    arrExclusions = ChargerExclusionsCalcul()
    If IsArray(arrExclusions) Then nbExclusions = UBound(arrExclusions, 1) Else nbExclusions = 0
    
    Dim rapport As String
    rapport = "=== DEBUG COLONNE " & colSelect & " ===" & vbCrLf
    rapport = rapport & "Jour: " & ws.Cells(ligneNumJour, colSelect).value & vbCrLf & vbCrLf
    
    Dim tot(1 To 11) As Double
    Dim totINF(1 To 11) As Double
    Dim i As Long, j As Long
    Dim cell As Range, code As String, nomPersonne As String
    Dim couleurCellule As Long, vals As Variant
    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
    Dim vLocal(1 To 11) As Double
    Dim codeFound As Boolean
    
    For j = 1 To 11: tot(j) = 0: totINF(j) = 0: Next j
    
    rapport = rapport & "CODES TROUVES:" & vbCrLf
    rapport = rapport & String(60, "-") & vbCrLf
    
    For i = ligneDebut To ligneFin
        Set cell = ws.Cells(i, colSelect)
        couleurCellule = cell.Interior.Color
        
        If couleurCellule <> couleurIgnore Then
            code = Trim(CStr(cell.value))
            
            If code <> "" Then
                nomPersonne = Trim(CStr(ws.Cells(i, 1).value))
                Dim fonctionPersonne As String: fonctionPersonne = ""
                
                If dictFonctions.Exists(Replace(nomPersonne, " ", "_")) Then
                    fonctionPersonne = dictFonctions(Replace(nomPersonne, " ", "_"))
                ElseIf dictFonctions.Exists(nomPersonne) Then
                    fonctionPersonne = dictFonctions(nomPersonne)
                End If
                
                ' Verifier si code d'absence (identique a la macro principale)
                Dim estAbsence As Boolean: estAbsence = False
                Dim codeUp As String: codeUp = UCase(code)
                If codeUp = "WE" Or codeUp Like "MAL*" Or codeUp Like "CA*" Or _
                   codeUp Like "RCT*" Or codeUp Like "MAT*" Or codeUp Like "MUT*" Or _
                   codeUp = "CTR" Or codeUp = "DP" Or codeUp = "RHS" Or _
                   codeUp = "EL" Or codeUp Like "AFC*" Or _
                   codeUp Like "3/4*" Or codeUp Like "4/5*" Or _
                   codeUp Like "CP*" Then
                    estAbsence = True
                End If
                
                For j = 1 To 11: vLocal(j) = 0: Next j
                codeFound = False
                
                ' 1) Essayer dictionnaire exact
                If dictCodes.Exists(code) Then
                    vals = dictCodes(code)
                    For j = 1 To 11: vLocal(j) = vals(j): Next j
                    codeFound = True
                    rapport = rapport & "L" & i & " [DICT] " & Left(nomPersonne, 15) & " : " & code
                    rapport = rapport & " => M=" & vLocal(1) & " PM=" & vLocal(2) & " S=" & vLocal(3) & " N=" & vLocal(4)
                    rapport = rapport & " | 645=" & vLocal(5) & " 7-8=" & vLocal(6) & " 1630=" & vLocal(7)
                    rapport = rapport & " | C15=" & vLocal(8) & " C20=" & vLocal(9) & " C20E=" & vLocal(10) & " C19=" & vLocal(11)
                    If fonctionPersonne <> "" Then rapport = rapport & " [FCT:" & fonctionPersonne & "]"
                    If estAbsence Then rapport = rapport & " [EXCLU:ABS]"
                    rapport = rapport & vbCrLf
                    
                ' 2) Sinon, parser le code comme horaire
                ElseIf ParseCode(code, h1, f1, h2, f2) Then
                    CalcPeriodes h1, f1, h2, f2, vLocal(1), vLocal(2), vLocal(3), vLocal(4)
                    CalcPresSpec h1, f1, h2, f2, vLocal(5), vLocal(6), vLocal(7)
                    If IsCodeC15(h1, f1, h2, f2) Then vLocal(8) = 1
                    If IsCodeC20(h1, f1, h2, f2) Then vLocal(9) = 1
                    If IsCodeC20E(h1, f1, h2, f2) Then vLocal(10) = 1
                    If IsCodeC19(h1, f1, h2, f2) Then vLocal(11) = 1
                    codeFound = True
                    rapport = rapport & "L" & i & " [PARSE] " & Left(nomPersonne, 15) & " : " & code
                    rapport = rapport & " => M=" & vLocal(1) & " PM=" & vLocal(2) & " S=" & vLocal(3) & " N=" & vLocal(4)
                    rapport = rapport & " | 645=" & vLocal(5) & " 7-8=" & vLocal(6) & " 1630=" & vLocal(7)
                    rapport = rapport & " | C15=" & vLocal(8) & " C20=" & vLocal(9) & " C20E=" & vLocal(10) & " C19=" & vLocal(11)
                    If fonctionPersonne <> "" Then rapport = rapport & " [FCT:" & fonctionPersonne & "]"
                    If estAbsence Then rapport = rapport & " [EXCLU:ABS]"
                    rapport = rapport & vbCrLf
                Else
                    rapport = rapport & "L" & i & " [???] " & Left(nomPersonne, 15) & " : " & code & " => NON RECONNU"
                    If estAbsence Then rapport = rapport & " [EXCLU:ABS]"
                    rapport = rapport & vbCrLf
                End If
                
                ' LOGIQUE DE FILTRAGE (IDENTIQUE A LA MACRO PRINCIPALE)
                Dim compterDansTotaux As Boolean: compterDansTotaux = False
                Dim fctUpper As String: fctUpper = UCase(fonctionPersonne)
                
                Dim estFonctionAutorisee As Boolean
                estFonctionAutorisee = (InStr("," & UCase(fonctionsACompter) & ",", "," & fctUpper & ",") > 0)
                
                If estFonctionAutorisee Then
                    If fctUpper = "CEFA" Then
                         Dim dictCEFAFormation As Object ' Chargement mini pour debug (ou on suppose inclus par defaut si pas acces)
                         ' Simplification debug: on compte le CEFA sauf si on veut etre tres precis (ignorer formation pour l'instant dans le debug ou ajouter le chargement)
                         ' Pour faire simple: on compte.
                         compterDansTotaux = True
                    Else
                        compterDansTotaux = True
                    End If
                End If
                
                ' Exclusions codes speciaux
                If estAbsence Then compterDansTotaux = False
                If fonctionPersonne = "" And Not estFonctionAutorisee Then compterDansTotaux = False
                
                ' Accumuler si eligible
                If codeFound And compterDansTotaux Then
                    For j = 1 To 11: tot(j) = tot(j) + vLocal(j): Next j
                    If UCase(fonctionPersonne) = "INF" Then
                        For j = 1 To 11: totINF(j) = totINF(j) + vLocal(j): Next j
                    End If
                End If
                
                ' Ajouter info "Compte/PasCompte" au rapport
                If codeFound Then
                    rapport = rapport & "   -> Compter: " & IIf(compterDansTotaux, "OUI", "NON") & " (Fct: " & fonctionPersonne & ")" & vbCrLf
                End If
            End If
        End If
    Next i
    
    rapport = rapport & vbCrLf & String(60, "-") & vbCrLf
    rapport = rapport & "TOTAUX CALCULES:" & vbCrLf
    rapport = rapport & "  Matin: " & tot(1) & " (INF: " & totINF(1) & ")" & vbCrLf
    rapport = rapport & "  PM: " & tot(2) & " (INF: " & totINF(2) & ")" & vbCrLf
    rapport = rapport & "  Soir: " & tot(3) & " (INF: " & totINF(3) & ")" & vbCrLf
    rapport = rapport & "  Nuit: " & tot(4) & " (INF: " & totINF(4) & ")" & vbCrLf
    rapport = rapport & "  P_0645: " & tot(5) & vbCrLf
    rapport = rapport & "  P_7H8H: " & tot(6) & vbCrLf
    rapport = rapport & "  P_1630: " & tot(7) & vbCrLf
    rapport = rapport & "  C15: " & tot(8) & vbCrLf
    rapport = rapport & "  C20: " & tot(9) & vbCrLf
    rapport = rapport & "  C20E: " & tot(10) & vbCrLf
    rapport = rapport & "  C19: " & tot(11) & vbCrLf
    
    rapport = rapport & vbCrLf & "CODES DANS DICTIONNAIRE: " & dictCodes.count
    
    ' Afficher dans la fenetre Immediate (Ctrl+G dans VBA)
    Debug.Print String(80, "=")
    Debug.Print rapport
    Debug.Print String(80, "=")
    
    MsgBox "Rapport affiche dans la fenetre Immediate (Ctrl+G)" & vbCrLf & vbCrLf & _
           "Totaux pour colonne " & colSelect & ":" & vbCrLf & _
           "Matin=" & tot(1) & " PM=" & tot(2) & " Soir=" & tot(3) & vbCrLf & _
           "P_0645=" & tot(5) & " P_7H8H=" & tot(6) & " P_1630=" & tot(7) & vbCrLf & _
           "C15=" & tot(8) & " C20=" & tot(9) & " C20E=" & tot(10) & " C19=" & tot(11), vbInformation
End Sub
