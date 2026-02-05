Attribute VB_Name = "Module_StylePremium"
Option Explicit

' =========================================================================================
'   MACRO STYLE PREMIUM - RELOOKING DU PLANNING
'   Objectif : Rendu moderne, lisible et professionnel (Flat Design)
'   Date : 21 janvier 2026
' =========================================================================================

Sub AppliquerStylePremium()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' --- 1. DEFINITION DE LA PALETTE (COULEURS MODERNES) ---
    Dim colFond As Long: colFond = RGB(255, 255, 255)       ' Blanc pur
    Dim colEnTete As Long: colEnTete = RGB(44, 62, 80)      ' Midnight Blue (En-tetes)
    Dim colTexteHeader As Long: colTexteHeader = RGB(236, 240, 241) ' Cloud White
    
    Dim colLignePair As Long: colLignePair = RGB(247, 249, 249) ' Gris tres pale
    Dim colLigneImpair As Long: colLigneImpair = RGB(255, 255, 255)
    
    Dim colMatin As Long: colMatin = RGB(52, 152, 219)      ' Peter River (Bleu)
    Dim colPM As Long: colPM = RGB(230, 126, 34)            ' Carrot (Orange)
    Dim colSoir As Long: colSoir = RGB(155, 89, 182)        ' Amethyst (Violet)
    Dim colNuit As Long: colNuit = RGB(52, 73, 94)          ' Wet Asphalt (Gris nuit)
    
    Dim colWeekEnd As Long: colWeekEnd = RGB(234, 237, 237) ' Gris weekend
    
    ' --- 2. CONFIGURATION GENERALE ---
    Application.ScreenUpdating = False
    
    With ws.Cells
        .Font.Name = "Segoe UI"  ' Police moderne
        .Font.Size = 9           ' Taille standard
    End With
    
    ' Masquer le quadrillage standard Excel
    ActiveWindow.DisplayGridlines = False
    
    ' --- 3. DETECTION DES ZONES (BasÃ© sur CalculFractionsPresence) ---
    ' On suppose que les totaux sont vers les lignes 60-70 comme vu precedemment
    Dim ligDebutTotaux As Long: ligDebutTotaux = 60
    Dim ligFinTotaux As Long: ligFinTotaux = 75 ' Large manœuvre
    Dim colDebutJours As Long: colDebutJours = 3
    Dim colFinJours As Long: colFinJours = 33
    
    ' --- 4. STYLE DE LA ZONE TOTAUX (Lignes 60+) ---
    Dim rngTotaux As Range
    Set rngTotaux = ws.Range(ws.Cells(ligDebutTotaux, 1), ws.Cells(ligFinTotaux, colFinJours))
    
    ' Nettoyage ancien style
    rngTotaux.Borders.LineStyle = xlNone
    rngTotaux.Interior.Color = xlNone
    
    ' Bordures fines horizontales (Zebra style)
    Dim i As Long
    For i = ligDebutTotaux To ligFinTotaux
        Dim r As Range: Set r = ws.Range(ws.Cells(i, 1), ws.Cells(i, colFinJours))
        
        ' Ligne de separation fine
        With r.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(220, 220, 220)
            .Weight = xlThin
        End With
        
        ' Alternance couleur de fond
        If i Mod 2 = 0 Then
            r.Interior.Color = colLignePair
        Else
            r.Interior.Color = colLigneImpair
        End If
        
        ' Alignement
        r.VerticalAlignment = xlCenter
        r.HorizontalAlignment = xlCenter
        
        ' Colonnes titres (A et B) : AlignÃ© Ã  gauche, gras
        With ws.Range(ws.Cells(i, 1), ws.Cells(i, 2))
            .HorizontalAlignment = xlLeft
            .Font.Bold = True
            .Font.Color = RGB(100, 100, 100)
            .IndentLevel = 1
        End With
    Next i
    
    ' --- 5. DATA BARS (BARRES DE DONNEES) POUR LES TOTAUX ---
    ' Matin (Ligne 60 approx) -> Barre Bleue
    Call AjouterDataBar(ws, 60, colDebutJours, colFinJours, colMatin)
    ' PM (Ligne 61 approx) -> Barre Orange
    Call AjouterDataBar(ws, 61, colDebutJours, colFinJours, colPM)
    ' Soir (Ligne 62 approx) -> Barre Violette
    Call AjouterDataBar(ws, 62, colDebutJours, colFinJours, colSoir)
    ' Nuit (Ligne 63 approx) -> Barre Gris Foncé
    Call AjouterDataBar(ws, 63, colDebutJours, colFinJours, colNuit)
    
    ' --- 6. EN-TETE DES JOURS (Lignes 3-4 approx) ---
    Dim rngHeader As Range
    Set rngHeader = ws.Range(ws.Cells(3, colDebutJours), ws.Cells(4, colFinJours))
    With rngHeader
        .Interior.Color = colEnTete
        .Font.Color = colTexteHeader
        .Font.Bold = True
        .Borders.LineStyle = xlNone
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Le style 'Premium' a ete applique avec succes !", vbInformation, "Design Updated"
End Sub

Private Sub AjouterDataBar(ws As Worksheet, lig As Long, cDebut As Long, cFin As Long, couleur As Long)
    ' Ajoute une barre de donnees conditionnelle sur la plage
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(lig, cDebut), ws.Cells(lig, cFin))
    
    ' Nettoyer formats conditionnels existants
    rng.FormatConditions.Delete
    
    ' Ajouter DataBar
    Dim db As Databar
    Set db = rng.FormatConditions.AddDatabar
    
    With db
        .ShowValue = True ' Garder le chiffre visible
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=15 ' Echelle fixe pour uniformite visuelle
        
        .BarColor.Color = couleur
        .BarFillType = xlDataBarFillSolid ' Solide = plus moderne que degrade
        .Direction = xlContext
        .AxisPosition = xlDataBarAxisAutomatic
        .BarBorder.Type = xlDataBarBorderNone
    End With
    On Error GoTo 0
End Sub
