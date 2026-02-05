' ExportedAt: 2026-01-13 15:05:00 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "GenerateurCalendrier_V2"
Option Explicit

'===================================================================================
' MODULE :      GenerateurCalendrier_V2 (Couleurs WE/Ferie separees)
' DESCRIPTION : Genere les en-tetes jours (Lun/Mar/...) + numeros (1..31),
'               recolore WE + jours feries avec couleurs separees, masque colonnes inutiles.
'
'               CONFIG-DRIVEN :
'               - Colonnes / lignes header via tblCFG
'               - Annee via CFG_Year
'               - Couleurs separees WE/Ferie via PLN_Couleur_* (format Long)
'
' PREREQUIS :   Module_Config importe (CfgLong / CfgText / CfgValue)
'===================================================================================

Public Sub GenererDatesEtJoursPourTousLesMois()
    Dim feuilleActuelle As Worksheet
    Dim dateJour As Date
    Dim annee As Long, indexMois As Integer
    Dim moisFrancais As Variant, nomsJours As Variant
    Dim wd As Integer, totalJours As Integer
    Dim jourFeries As Collection
    Dim i As Long, col As Long
    Dim isHoliday As Boolean
    
    ' --- config-driven layout ---
    Dim FIRST_DAY_COL As Long
    Dim LAST_DAY_COL As Long
    Dim ROW_JOUR_SEMAINE As Long
    Dim ROW_NUMERO_JOUR As Long
    
    Dim keepHeaderRows As String
    Dim yearCellAddress As String
    Dim localeKey As String
    
    ' --- config-driven colors (separees WE/Ferie) ---
    Dim colorWorkday As Long
    Dim colorWeekend As Long
    Dim colorFerie As Long
    Dim colorPoliceWE As Long
    Dim colorPoliceFerie As Long
    
    On Error GoTo EH
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '========================
    ' Lire configuration
    '========================
    FIRST_DAY_COL = CLng(CfgValueOr("PLN_FirstDayCol", 0))
    LAST_DAY_COL = CLng(CfgValueOr("PLN_LastDayCol", 0))
    ROW_JOUR_SEMAINE = CLng(CfgValueOr("PLN_Row_DayNames", 0))
    ROW_NUMERO_JOUR = CLng(CfgValueOr("PLN_Row_DayNumbers", 0))
    
    annee = CLng(CfgValueOr("CFG_Year", 0))
    
    keepHeaderRows = CfgTextOr("VIEW_HeaderRows_Keep", "")
    yearCellAddress = CfgTextOr("VIEW_YearCell", "")
    If Len(yearCellAddress) = 0 Then yearCellAddress = "B1"
    
    localeKey = CfgTextOr("CFG_Locale", "")
    
    ' Couleurs jours normaux
    colorWorkday = ParseRgbTextToLong(CfgTextOr("PAL_Color_Workday", ""), RGB(204, 229, 255))
    
    ' Couleurs WE et Feries separees (format Long)
    colorWeekend = CLng(CfgValueOr("PLN_Couleur_Weekend", 15773696))
    colorFerie = CLng(CfgValueOr("PLN_Couleur_Ferie", 255))
    colorPoliceWE = CLng(CfgValueOr("PLN_Couleur_Police_Weekend", 16777215))
    colorPoliceFerie = CLng(CfgValueOr("PLN_Couleur_Police_Ferie", 16777215))
    
    '========================
    ' Validations
    '========================
    If annee < 1900 Or annee > 2100 Then
        MsgBox "Annee non valide (CFG_Year).", vbCritical
        GoTo CleanUp
    End If
    
    If FIRST_DAY_COL <= 0 Or LAST_DAY_COL <= 0 Or LAST_DAY_COL < FIRST_DAY_COL Then
        MsgBox "Configuration colonnes invalide.", vbCritical
        GoTo CleanUp
    End If
    
    '========================
    ' Libelles mois / jours
    '========================
    moisFrancais = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                         "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    
    If LCase$(localeKey) = "en-us" Or LCase$(localeKey) = "en-gb" Then
        nomsJours = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    Else
        nomsJours = Array("Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim")
    End If
    
    '========================
    ' Jours feries (BE)
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
            
            ' 1) Afficher annee
            On Error Resume Next
            With feuilleActuelle.Range(yearCellAddress)
                .Value = annee
                .Font.Bold = True
            End With
            On Error GoTo EH
            
            ' 2) Reafficher colonnes
            feuilleActuelle.Range( _
                feuilleActuelle.Cells(1, FIRST_DAY_COL), _
                feuilleActuelle.Cells(1, LAST_DAY_COL) _
            ).EntireColumn.Hidden = False
            
            ' 3) Nettoyer ancien en-tete
            With feuilleActuelle.Range( _
                feuilleActuelle.Cells(ROW_JOUR_SEMAINE, FIRST_DAY_COL), _
                feuilleActuelle.Cells(ROW_NUMERO_JOUR, LAST_DAY_COL))
                .ClearContents
            End With
            ClearHeaderFill feuilleActuelle, ROW_JOUR_SEMAINE, ROW_NUMERO_JOUR, FIRST_DAY_COL, LAST_DAY_COL
            
            ' 4) Nombre de jours du mois
            totalJours = Day(DateSerial(annee, indexMois + 1, 0))
            
            ' 5) Preparer tableau
            Dim arrHeaders() As Variant
            ReDim arrHeaders(1 To 2, 1 To totalJours)
            
            For i = 1 To totalJours
                dateJour = DateSerial(annee, indexMois, i)
                wd = Weekday(dateJour, vbMonday)
                arrHeaders(1, i) = nomsJours(wd - 1)
                arrHeaders(2, i) = i
            Next i
            
            ' 6) Ecrire valeurs
            feuilleActuelle.Cells(ROW_JOUR_SEMAINE, FIRST_DAY_COL) _
                .Resize(2, totalJours).Value = arrHeaders
            
            ' 7) Recolorier chaque colonne (WE vs Ferie vs Normal)
            Dim rngHeader As Range
            For col = FIRST_DAY_COL To FIRST_DAY_COL + totalJours - 1
                
                dateJour = DateSerial(annee, indexMois, feuilleActuelle.Cells(ROW_NUMERO_JOUR, col).Value)
                wd = Weekday(dateJour, vbMonday)
                isHoliday = EstJourFerie(dateJour, jourFeries)
                
                Set rngHeader = feuilleActuelle.Range( _
                    feuilleActuelle.Cells(ROW_JOUR_SEMAINE, col), _
                    feuilleActuelle.Cells(ROW_NUMERO_JOUR, col))
                
                If isHoliday Then
                    ' Jour ferie = fond rouge, police blanche
                    rngHeader.Interior.Color = colorFerie
                    rngHeader.Font.Color = colorPoliceFerie
                    rngHeader.Font.Bold = True
                ElseIf wd >= 6 Then
                    ' Weekend = fond bleu clair, police blanche
                    rngHeader.Interior.Color = colorWeekend
                    rngHeader.Font.Color = colorPoliceWE
                    rngHeader.Font.Bold = True
                Else
                    ' Jour normal
                    rngHeader.Interior.Color = colorWorkday
                    rngHeader.Font.Color = 0
                    rngHeader.Font.Bold = False
                End If
            Next col
            
            ' 8) Masquer colonnes apres dernier jour
            If totalJours < 31 Then
                feuilleActuelle.Range( _
                    feuilleActuelle.Cells(1, FIRST_DAY_COL + totalJours), _
                    feuilleActuelle.Cells(1, LAST_DAY_COL) _
                ).EntireColumn.Hidden = True
            End If
            
            ' 9) Forcer lignes en-tete visibles
            If Len(keepHeaderRows) > 0 Then
                feuilleActuelle.Rows(keepHeaderRows).Hidden = False
            End If
            
        End If
    Next indexMois
    
    MsgBox "Calendriers generes pour " & annee & " (WE bleu, Feries rouge).", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

EH:
    MsgBox "Erreur: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanUp
End Sub


'===================================================================================
' UTILS
'===================================================================================

Private Function ParseRgbTextToLong(ByVal rgbText As String, ByVal defaultColor As Long) As Long
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
        .Interior.Pattern = xlNone
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
    feries.Add DateSerial(annee, 1, 1), CStr(DateSerial(annee, 1, 1))
    feries.Add paques + 1, CStr(paques + 1)
    feries.Add DateSerial(annee, 5, 1), CStr(DateSerial(annee, 5, 1))
    feries.Add paques + 39, CStr(paques + 39)
    feries.Add paques + 50, CStr(paques + 50)
    feries.Add DateSerial(annee, 7, 21), CStr(DateSerial(annee, 7, 21))
    feries.Add DateSerial(annee, 8, 15), CStr(DateSerial(annee, 8, 15))
    feries.Add DateSerial(annee, 11, 1), CStr(DateSerial(annee, 11, 1))
    feries.Add DateSerial(annee, 11, 11), CStr(DateSerial(annee, 11, 11))
    feries.Add DateSerial(annee, 12, 25), CStr(DateSerial(annee, 12, 25))
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
