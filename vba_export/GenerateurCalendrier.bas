Attribute VB_Name = "GenerateurCalendrier"
' ExportedAt: 2026-01-12 15:37:08 | Workbook: Planning_2026.xlsm
Option Explicit

'===================================================================================
' MODULE :      GenerateurCalendrier (FINAL - Config-Driven + corrigé colonne C)
' DESCRIPTION : Génère les en-têtes jours (Lun/Mar/...) + numéros (1..31),
'               recolore WE + jours fériés, masque colonnes inutiles.
'
'               CONFIG-DRIVEN :
'               - Colonnes / lignes header via tblCFG
'               - Année via CFG_Year
'               - Cellule année via VIEW_YearCell (fallback B1)
'               - Lignes à garder visibles via VIEW_HeaderRows_Keep
'               - Couleurs via PAL_Color_Workday / PAL_Color_WeekendOrHoliday (fallback RGB)
'
' PREREQUIS :   Module_Config importé (CfgLong / CfgText)
'===================================================================================

Public Sub GenererDatesEtJoursPourTousLesMois()
    Dim feuilleActuelle As Worksheet
    Dim dateJour As Date
    Dim annee As Long, indexMois As Integer
    Dim moisFrancais As Variant, nomsJours As Variant
    Dim wd As Integer, totalJours As Integer
    Dim jourFeries As Collection
    Dim i As Long, col As Long
    
    ' --- config-driven layout ---
    Dim FIRST_DAY_COL As Long
    Dim LAST_DAY_COL As Long
    Dim ROW_JOUR_SEMAINE As Long
    Dim ROW_NUMERO_JOUR As Long
    
    Dim keepHeaderRows As String
    Dim yearCellAddress As String
    Dim localeKey As String
    
    ' --- config-driven colors ---
    Dim colorWorkday As Long
    Dim colorWeekendOrHoliday As Long
    
    On Error GoTo EH
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '========================
    ' Lire configuration
    '========================
    FIRST_DAY_COL = CLng(Module_Config.CfgValueOr("PLN_FirstDayCol", 0))          ' ex: 3 (C)
    LAST_DAY_COL = CLng(Module_Config.CfgValueOr("PLN_LastDayCol", 0))            ' ex: 33 (AG)
    ROW_JOUR_SEMAINE = CLng(Module_Config.CfgValueOr("PLN_Row_DayNames", 0))      ' ex: 3
    ROW_NUMERO_JOUR = CLng(Module_Config.CfgValueOr("PLN_Row_DayNumbers", 0))     ' ex: 4
    
    annee = CLng(Module_Config.CfgValueOr("CFG_Year", 0))                         ' ex: 2026
    
    keepHeaderRows = Module_Config.CfgTextOr("VIEW_HeaderRows_Keep", "")    ' ex: "3:4" ou "1:4"
    yearCellAddress = Module_Config.CfgTextOr("VIEW_YearCell", "")          ' ex: "B1" (fallback si vide)
    If Len(yearCellAddress) = 0 Then yearCellAddress = "B1"
    
    localeKey = Module_Config.CfgTextOr("CFG_Locale", "")                   ' ex: "fr-BE"
    
    ' Couleurs (fallback sur valeurs historiques si clé absente / invalide)
    colorWorkday = ParseRgbTextToLong(Module_Config.CfgTextOr("PAL_Color_Workday", ""), RGB(204, 229, 255))
    colorWeekendOrHoliday = ParseRgbTextToLong(Module_Config.CfgTextOr("PAL_Color_WeekendOrHoliday", ""), RGB(255, 0, 0))
    
    '========================
    ' Validations (safe)
    '========================
    If annee < 1900 Or annee > 2100 Then
        MsgBox "Année non valide (CFG_Year).", vbCritical
        GoTo CleanUp
    End If
    
    If FIRST_DAY_COL <= 0 Or LAST_DAY_COL <= 0 Or LAST_DAY_COL < FIRST_DAY_COL Then
        MsgBox "Configuration colonnes invalide (PLN_FirstDayCol / PLN_LastDayCol).", vbCritical
        GoTo CleanUp
    End If
    
    If ROW_JOUR_SEMAINE <= 0 Or ROW_NUMERO_JOUR <= 0 Or ROW_NUMERO_JOUR < ROW_JOUR_SEMAINE Then
        MsgBox "Configuration lignes invalide (PLN_Row_DayNames / PLN_Row_DayNumbers).", vbCritical
        GoTo CleanUp
    End If
    
    '========================
    ' Libellés mois / jours
    '========================
    moisFrancais = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                         "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    
    ' Locale minimal (fallback fr si vide)
    If LCase$(localeKey) = "en-us" Or LCase$(localeKey) = "en-gb" Then
        nomsJours = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    Else
        nomsJours = Array("Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim")
    End If
    
    '========================
    ' Jours fériés (BE)
    '========================
    Set jourFeries = BuildFeriesBE(annee)
    
    '========================
    ' Boucle 12 mois
    '========================
    For indexMois = 1 To 12
        
        Set feuilleActuelle = Nothing
        On Error Resume Next
        Set feuilleActuelle = Sheets(moisFrancais(indexMois - 1))
        On Error GoTo EH
        
        If Not feuilleActuelle Is Nothing Then
            
            ' 1) Afficher l'année dans la cellule config-driven (VIEW_YearCell)
            '    Safe: si adresse invalide, ne crash pas
            On Error Resume Next
            With feuilleActuelle.Range(yearCellAddress)
                .value = annee
                .Font.Bold = True
            End With
            On Error GoTo EH
            
            ' 2) Réafficher colonnes zone jours au cas où
            feuilleActuelle.Range( _
                feuilleActuelle.Cells(1, FIRST_DAY_COL), _
                feuilleActuelle.Cells(1, LAST_DAY_COL) _
            ).EntireColumn.Hidden = False
            
            ' 3) Nettoyer ancien en-tête (jours + numéros + couleurs)
            With feuilleActuelle.Range( _
                feuilleActuelle.Cells(ROW_JOUR_SEMAINE, FIRST_DAY_COL), _
                feuilleActuelle.Cells(ROW_NUMERO_JOUR, LAST_DAY_COL))
                .ClearContents
            End With
            ClearHeaderFill feuilleActuelle, ROW_JOUR_SEMAINE, ROW_NUMERO_JOUR, FIRST_DAY_COL, LAST_DAY_COL
            
            ' 4) Nombre de jours du mois
            totalJours = Day(DateSerial(annee, indexMois + 1, 0))
            
            ' 5) Préparer tableau 2 lignes : jour semaine / numéro
            Dim arrHeaders() As Variant
            ReDim arrHeaders(1 To 2, 1 To totalJours)
            
            For i = 1 To totalJours
                dateJour = DateSerial(annee, indexMois, i)
                wd = Weekday(dateJour, vbMonday) ' 1=Lun ... 7=Dim
                
                arrHeaders(1, i) = nomsJours(wd - 1)
                arrHeaders(2, i) = i
            Next i
            
            ' 6) Écrire valeurs à partir de (ROW_JOUR_SEMAINE, FIRST_DAY_COL)
            feuilleActuelle.Cells(ROW_JOUR_SEMAINE, FIRST_DAY_COL) _
                .Resize(2, totalJours).value = arrHeaders
            
            ' 7) Recolorier chaque colonne jour par jour
            For col = FIRST_DAY_COL To FIRST_DAY_COL + totalJours - 1
                
                dateJour = DateSerial(annee, indexMois, feuilleActuelle.Cells(ROW_NUMERO_JOUR, col).value)
                wd = Weekday(dateJour, vbMonday)
                
                If wd >= 6 Or EstJourFerie(dateJour, jourFeries) Then
                    feuilleActuelle.Range( _
                        feuilleActuelle.Cells(ROW_JOUR_SEMAINE, col), _
                        feuilleActuelle.Cells(ROW_NUMERO_JOUR, col) _
                    ).Interior.Color = colorWeekendOrHoliday
                Else
                    feuilleActuelle.Range( _
                        feuilleActuelle.Cells(ROW_JOUR_SEMAINE, col), _
                        feuilleActuelle.Cells(ROW_NUMERO_JOUR, col) _
                    ).Interior.Color = colorWorkday
                End If
            Next col
            
            ' 8) Masquer colonnes après dernier jour du mois
            If totalJours < 31 Then
                feuilleActuelle.Range( _
                    feuilleActuelle.Cells(1, FIRST_DAY_COL + totalJours), _
                    feuilleActuelle.Cells(1, LAST_DAY_COL) _
                ).EntireColumn.Hidden = True
            End If
            
            ' 9) Forcer les lignes d'en-tête visibles après génération
            If Len(keepHeaderRows) > 0 Then
                feuilleActuelle.Rows(keepHeaderRows).Hidden = False
            End If
            
        End If
    Next indexMois
    
    MsgBox "Calendriers générés pour l'année " & annee & " (config-driven).", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

EH:
    MsgBox "Erreur GenererDatesEtJoursPourTousLesMois: " & Err.Number & " - " & Err.description, vbCritical
    Resume CleanUp
End Sub


'===================================================================================
' UTILS - Colors
'===================================================================================

Private Function ParseRgbTextToLong(ByVal rgbText As String, ByVal defaultColor As Long) As Long
    ' Attend "R,G,B" (ex: "204,229,255"). Retourne un Long compatible .Interior.Color
    Dim parts() As String
    Dim r As Long, g As Long, b As Long
    
    rgbText = Trim$(rgbText)
    If Len(rgbText) = 0 Then
        ParseRgbTextToLong = defaultColor
        Exit Function
    End If
    
    parts = Split(rgbText, ",")
    If UBound(parts) <> 2 Then
        ParseRgbTextToLong = defaultColor
        Exit Function
    End If
    
    On Error GoTo SafeFail
    r = CLng(Trim$(parts(0)))
    g = CLng(Trim$(parts(1)))
    b = CLng(Trim$(parts(2)))
    
    If r < 0 Or r > 255 Or g < 0 Or g > 255 Or b < 0 Or b > 255 Then GoTo SafeFail
    
    ParseRgbTextToLong = RGB(r, g, b)
    Exit Function
    
SafeFail:
    ParseRgbTextToLong = defaultColor
End Function

Private Sub ClearHeaderFill(ByVal ws As Worksheet, ByVal rowDayNames As Long, ByVal rowDayNumbers As Long, ByVal firstCol As Long, ByVal lastCol As Long)
    With ws.Range(ws.Cells(rowDayNames, firstCol), ws.Cells(rowDayNumbers, lastCol))
        .Interior.pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
End Sub


'===================================================================================
' JOURS FERIES (BE)
'===================================================================================

Private Function BuildFeriesBE(ByVal annee As Long) As Collection
    Dim feries As New Collection
    Dim paques As Date
    
    paques = CalculerPaques(annee)
    
    On Error Resume Next
    feries.Add DateSerial(annee, 1, 1), CStr(DateSerial(annee, 1, 1))            ' Nouvel An
    feries.Add paques + 1, CStr(paques + 1)                                      ' Lundi de Pâques
    feries.Add DateSerial(annee, 5, 1), CStr(DateSerial(annee, 5, 1))            ' Fête du travail
    feries.Add paques + 39, CStr(paques + 39)                                    ' Ascension
    feries.Add paques + 50, CStr(paques + 50)                                    ' Lundi de Pentecôte
    feries.Add DateSerial(annee, 7, 21), CStr(DateSerial(annee, 7, 21))          ' Fête nationale (BE)
    feries.Add DateSerial(annee, 8, 15), CStr(DateSerial(annee, 8, 15))          ' Assomption
    feries.Add DateSerial(annee, 11, 1), CStr(DateSerial(annee, 11, 1))          ' Toussaint
    feries.Add DateSerial(annee, 11, 11), CStr(DateSerial(annee, 11, 11))        ' Armistice
    feries.Add DateSerial(annee, 12, 25), CStr(DateSerial(annee, 12, 25))        ' Noël
    On Error GoTo 0
    
    Set BuildFeriesBE = feries
End Function

Private Function EstJourFerie(ByVal d As Date, ByVal feries As Collection) As Boolean
    On Error Resume Next
    Dim tmp As Variant
    tmp = feries(CStr(d))
    EstJourFerie = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function CalculerPaques(ByVal annee As Long) As Date
    Dim a As Integer, b As Integer, c As Integer
    Dim d As Integer, e As Integer, f As Integer
    Dim g As Integer, h As Integer, i As Integer
    Dim k As Integer, l As Integer, m As Integer
    Dim mois As Integer, jour As Integer
    
    a = annee Mod 19
    b = annee \ 100
    c = annee Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    mois = (h + l - 7 * m + 114) \ 31
    jour = ((h + l - 7 * m + 114) Mod 31) + 1
    
    CalculerPaques = DateSerial(annee, mois, jour)
End Function




'===================================================================================
' MAINTENANCE: Corriger #REF! des feuilles mensuelles
' - A1 -> =Feuil_Config!$B$27
'===================================================================================
Public Sub FixREFErrors_MonthSheets()
    Dim months As Variant
    months = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                   "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")

    Dim m As Variant, ws As Worksheet
    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    For Each m In months
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(m))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Range("A1").Formula = "=Feuil_Config!$B$27"
        End If
    Next m

CleanUp:
    Application.ScreenUpdating = True
End Sub
