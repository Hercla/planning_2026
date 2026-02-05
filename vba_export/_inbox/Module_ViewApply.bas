' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module_ViewApply"
Option Explicit

'===================================================================================
' MODULE: Module_ViewApply (MAJ - Local-First + Scope + Safe-Guards)
' PURPOSE:
'   Appliquer les réglages VIEW depuis tblCFG :
'   - Zoom
'   - Masquage colonnes menu
'   - Masquage colonne B
'   - Auto-hide des lignes "noms" si vide (col A ou autre)
'   - Masquages JOUR/NUIT via listes de blocs ("5:5;31:39;...")
'
'   + MODE LOCAL PAR FEUILLE (LOCAL-FIRST)
'   - Toggle JOUR/NUIT sur l'onglet actif seulement
'   - Stockage du mode local via Name de feuille : TEAM_MODE_LOCAL
'
'   + SCOPE CFG (VIEW_ApplyScope = ACTIVE/ALL)
'   - Routeur VIEW_Apply_ByScope
'
'   + SAFE-GUARDS
'   - Ne jamais appliquer la vue sur Feuil_Config / feuilles non-mois
'   - Empêche le masquage accidentel de la colonne B (Valeur) sur la config
'
' PREREQUIS:
'   Module_Config (CfgLong/CfgText/CfgBool, ou CfgBool via Module_Config)
'===================================================================================

'===================================================================================
' ROUTEUR SCOPE
'===================================================================================

Public Sub VIEW_Apply_ByScope()
    Dim scope As String
    scope = UCase$(Trim$(CfgTextOr("VIEW_ApplyScope", ""))) ' ACTIVE / ALL
    
    If scope = "ALL" Then
        VIEW_ApplyToAllMonthSheets
    Else
        VIEW_ApplyToActiveSheet
    End If
End Sub

'===================================================================================
' APIS PUBLIQUES
'===================================================================================

Public Sub VIEW_ApplyToAllMonthSheets()
    Dim m As Variant
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    On Error GoTo CleanUp
    
    For Each m In MonthNamesArray()
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(m))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            VIEW_ApplyToSheet ws
        End If
    Next m

CleanUp:
    Application.ScreenUpdating = True
End Sub

Public Sub VIEW_ApplyToActiveSheet()
    If ActiveSheet Is Nothing Then Exit Sub
    VIEW_ApplyToSheet ActiveSheet
End Sub

' Applique la vue sur une feuille en respectant la logique:
' 1) TEAM_MODE_LOCAL si présent sur la feuille
' 2) sinon TEAM_MODE global depuis tblCFG
Public Sub VIEW_ApplyToSheet(ByVal ws As Worksheet)
    Dim modeLocal As String
    Dim teamMode As String
    
    '========================
    ' SAFE-GUARDS (critique)
    '========================
    If Not ShouldApplyViewToSheet(ws) Then Exit Sub
    
    '========================
    ' MODE (local-first)
    '========================
    modeLocal = GetLocalMode(ws)
    If modeLocal = "JOUR" Or modeLocal = "NUIT" Then
        teamMode = modeLocal
    Else
        teamMode = UCase$(Trim$(CfgTextOr("TEAM_MODE", ""))) ' fallback global
        If teamMode <> "NUIT" Then teamMode = "JOUR"
    End If
    
    VIEW_ApplyToSheet_WithMode ws, teamMode
End Sub

'===================================================================================
' MODE LOCAL (par feuille)
'===================================================================================

' Force le mode sur la feuille active seulement + mémorise
Public Sub VIEW_SetMode_ActiveSheet(ByVal mode As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If Not ShouldApplyViewToSheet(ws) Then Exit Sub
    
    mode = UCase$(Trim$(mode))
    If mode <> "JOUR" And mode <> "NUIT" Then Exit Sub
    
    SetLocalMode ws, mode
    VIEW_ApplyToSheet_WithMode ws, mode
End Sub

' Toggle JOUR/NUIT sur la feuille active seulement + mémorise
Public Sub VIEW_ToggleMode_ActiveSheet()
    Dim ws As Worksheet
    Dim cur As String
    
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If Not ShouldApplyViewToSheet(ws) Then Exit Sub
    
    cur = GetLocalMode(ws)
    If cur <> "JOUR" And cur <> "NUIT" Then
        cur = UCase$(Trim$(CfgTextOr("TEAM_MODE", "")))
        If cur <> "NUIT" Then cur = "JOUR"
    End If
    
    If cur = "NUIT" Then
        VIEW_SetMode_ActiveSheet "JOUR"
    Else
        VIEW_SetMode_ActiveSheet "NUIT"
    End If
End Sub

' Supprime le mode local de la feuille active (retour au global TEAM_MODE)
Public Sub VIEW_ClearLocalMode_ActiveSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If Not ShouldApplyViewToSheet(ws) Then Exit Sub
    
    ClearLocalMode ws
    VIEW_ApplyToSheet ws
End Sub

'===================================================================================
' ORCHESTRATEURS
'===================================================================================

' Global: Génère calendrier + applique vue selon scope
Public Sub RUN_Calendar_And_View()
    GenerateurCalendrier.GenererDatesEtJoursPourTousLesMois
    VIEW_Apply_ByScope
End Sub

'===================================================================================
' APPLY CORE (avec mode explicite)
'===================================================================================

Private Sub VIEW_ApplyToSheet_WithMode(ByVal ws As Worksheet, ByVal teamMode As String)
    ApplyZoom ws
    ApplyHideMenuCols ws
    ApplyHideColumnB ws
    ApplyAutoHideNames ws
    
    If teamMode = "NUIT" Then
        ApplyHideBlocks ws, CfgTextOr("VIEW_Nuit_HideBlocks", "")
    Else
        ApplyHideBlocks ws, CfgTextOr("VIEW_Jour_HideBlocks", "")
    End If
    
    ' Forcer header visible (si configuré)
    Dim keepHeaderRows As String
    keepHeaderRows = CfgTextOr("VIEW_HeaderRows_Keep", "")
    If Len(keepHeaderRows) > 0 Then ws.Rows(keepHeaderRows).Hidden = False
End Sub

'----------------------------
' Zoom
'----------------------------
Private Sub ApplyZoom(ByVal ws As Worksheet)
    Dim z As Long
    z = CfgLong("VIEW_Zoom")
    If z < 10 Or z > 400 Then Exit Sub
    
    On Error Resume Next
    ws.Activate
    ActiveWindow.Zoom = z
    On Error GoTo 0
End Sub

'----------------------------
' Colonnes menu à masquer
'----------------------------
Private Sub ApplyHideMenuCols(ByVal ws As Worksheet)
    Dim cols As String
    cols = Trim$(CfgTextOr("VIEW_MenuCols", "")) ' ex: "AH:AO"
    If Len(cols) = 0 Then Exit Sub
    
    On Error Resume Next
    ws.Columns(cols).Hidden = True
    On Error GoTo 0
End Sub

'----------------------------
' Masquer colonne B si demandé
'----------------------------
Private Sub ApplyHideColumnB(ByVal ws As Worksheet)
    ' Safe: ce code ne s'exécute pas sur Feuil_Config grâce à ShouldApplyViewToSheet
    If CfgBool("VIEW_HideColumnB") Then
        ws.Columns("B").Hidden = True
    Else
        ws.Columns("B").Hidden = False
    End If
End Sub

'----------------------------
' Auto-hide lignes "noms" si cellule vide (col A ou autre)
'----------------------------
Private Sub ApplyAutoHideNames(ByVal ws As Worksheet)
    Dim colLetter As String
    Dim firstRow As Long, lastRow As Long
    Dim r As Long
    
    colLetter = Trim$(CfgTextOr("VIEW_NameCol_A", "")) ' ex: "A"
    If Len(colLetter) = 0 Then colLetter = "A"
    
    firstRow = CfgLong("VIEW_AutoHide_FirstRow")
    lastRow = CfgLong("VIEW_AutoHide_LastRow")
    If firstRow <= 0 Or lastRow <= 0 Or lastRow < firstRow Then Exit Sub
    
    For r = firstRow To lastRow
        If Len(Trim$(CStr(ws.Range(colLetter & r).value))) = 0 Then
            ws.Rows(r).Hidden = True
        Else
            ws.Rows(r).Hidden = False
        End If
    Next r
End Sub

'----------------------------
' Appliquer les blocs "a:b;c:d;..."
'----------------------------
Private Sub ApplyHideBlocks(ByVal ws As Worksheet, ByVal blocks As String)
    Dim items() As String, it As Variant
    Dim p() As String
    Dim a As Long, b As Long
    
    blocks = Trim$(blocks)
    If Len(blocks) = 0 Then Exit Sub
    
    items = Split(blocks, ";")
    For Each it In items
        If Len(Trim$(CStr(it))) > 0 Then
            p = Split(Trim$(CStr(it)), ":")
            If UBound(p) = 1 Then
                a = val(p(0))
                b = val(p(1))
                If a > 0 And b > 0 And b >= a Then
                    ws.Rows(a & ":" & b).Hidden = True
                End If
            End If
        End If
    Next it
End Sub

'===================================================================================
' SAFE-GUARDS
'===================================================================================

Private Function ShouldApplyViewToSheet(ByVal ws As Worksheet) As Boolean
    Dim cfgSheetName As String
    
    ' 1) Ne jamais appliquer sur la feuille config (sinon tu masques "Valeur" en col B)
    cfgSheetName = CfgTextOr("SHEET_FeuilConfig", "")
    If Len(cfgSheetName) = 0 Then cfgSheetName = "Feuil_Config"
    
    If StrComp(ws.Name, cfgSheetName, vbTextCompare) = 0 Then
        ShouldApplyViewToSheet = False
        Exit Function
    End If
    
    If StrComp(ws.Name, "Feuil_Config", vbTextCompare) = 0 Then
        ShouldApplyViewToSheet = False
        Exit Function
    End If
    
    ' 2) Par défaut: appliquer uniquement sur les feuilles mois
    If Not IsMonthSheet(ws.Name) Then
        ShouldApplyViewToSheet = False
        Exit Function
    End If
    
    ShouldApplyViewToSheet = True
End Function

Private Function IsMonthSheet(ByVal sheetName As String) As Boolean
    Dim m As Variant
    For Each m In MonthNamesArray()
        If StrComp(sheetName, CStr(m), vbTextCompare) = 0 Then
            IsMonthSheet = True
            Exit Function
        End If
    Next m
    IsMonthSheet = False
End Function

Private Function MonthNamesArray() As Variant
    MonthNamesArray = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                            "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
End Function

'===================================================================================
' STORAGE: Mode local via Name scoped à la feuille (TEAM_MODE_LOCAL)
'===================================================================================

Private Sub SetLocalMode(ByVal ws As Worksheet, ByVal mode As String)
    Dim nm As Name
    Dim nmFull As String
    
    nmFull = ws.Name & "!TEAM_MODE_LOCAL"
    
    On Error Resume Next
    Set nm = ThisWorkbook.names(nmFull)
    On Error GoTo 0
    
    If nm Is Nothing Then
        ws.names.Add Name:="TEAM_MODE_LOCAL", RefersTo:="=""" & mode & """"
    Else
        nm.RefersTo = "=""" & mode & """"
    End If
End Sub

Private Function GetLocalMode(ByVal ws As Worksheet) As String
    Dim nm As Name
    Dim nmFull As String
    Dim s As String
    
    nmFull = ws.Name & "!TEAM_MODE_LOCAL"
    
    On Error Resume Next
    Set nm = ThisWorkbook.names(nmFull)
    On Error GoTo 0
    
    If nm Is Nothing Then
        GetLocalMode = ""
        Exit Function
    End If
    
    s = nm.RefersTo          ' ex: ="JOUR"
    s = Replace(s, "=", "")
    s = Replace(s, """", "")
    GetLocalMode = UCase$(Trim$(s))
End Function

Private Sub ClearLocalMode(ByVal ws As Worksheet)
    Dim nm As Name
    Dim nmFull As String
    
    nmFull = ws.Name & "!TEAM_MODE_LOCAL"
    
    On Error Resume Next
    Set nm = ThisWorkbook.names(nmFull)
    If Not nm Is Nothing Then nm.Delete
    On Error GoTo 0
End Sub


