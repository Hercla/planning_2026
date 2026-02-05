Attribute VB_Name = "ModuleModes"
Option Explicit

' =====================================================================================
'   MACRO DE GESTION D'AFFICHAGE POUR PLANNING
' =====================================================================================
'   Auteur:          Adapté et amélioré
'   Date:            27 juin 2025
'   Description:     Bascule entre un affichage "Jour" et "Nuit" pour un planning,
'                    en masquant/affichant dynamiquement les lignes pertinentes.
' =====================================================================================


' --- Énumération pour rendre le code plus lisible ---
Private Enum ViewMode
    ViewJour
    ViewNuit
End Enum


' --- Constantes globales pour une maintenance facile ---
Private Const START_SCHEDULE_COL As String = "B"
Private Const END_SCHEDULE_COL As String = "AG"
Private Const MENU_COLS As String = "AH:AO"
Private Const NAME_RANGE_TO_CHECK As String = "A1:A50" ' Plage large pour sécurité


' =====================================================================================
'   MACROS PUBLIQUES (À AFFECTER AUX BOUTONS SUR LA FEUILLE EXCEL)
' =====================================================================================

Public Sub Mode_Jour_Ancien()
    AdjustView ViewJour
End Sub

Public Sub Mode_Nuit_Ancien()
    AdjustView ViewNuit
End Sub


' =====================================================================================
'   FONCTION CENTRALE (PRIVÉE)
' =====================================================================================
Private Sub AdjustView(mode As ViewMode)
    Dim ws As Worksheet
    Dim dynamicRows As Variant
    Dim rowsToHide As Variant

    Set ws = ActiveSheet

    ' --- Optimisation de la performance ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo ErrorHandler

    ' 1. RÉINITIALISATION : On commence par tout afficher pour partir d'un état propre.
    ws.Cells.EntireRow.Hidden = False
    ws.Cells.EntireColumn.Hidden = False
    
    ' On ré-affiche également toutes les formes de commentaires au cas où
    Dim cmt As Comment
    On Error Resume Next
    For Each cmt In ws.Comments
        cmt.Shape.Visible = True
    Next cmt
    On Error GoTo 0


    ' 2. DÉFINITION DES RÈGLES D'AFFICHAGE
    If mode = ViewJour Then
        ' Mode Jour :
        ' On vérifie dynamiquement le personnel de Jour (6-28) et les remplacements Jour (40-42).
        dynamicRows = Array(6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 40, 41, 42)
        
        ' On masque de force le personnel de Nuit (31-39), les remplacements Nuit (46-47) et les totaux de Nuit.
        rowsToHide = Array("5:5", "31:39", "43:45", "46:47", "48:58", "71:150")

    Else ' Mode Nuit
        ' ============================================================================================
        ' ### RÈGLES STRICTES APPLIQUÉES POUR LE MODE NUIT ###
        ' ============================================================================================
        
        ' Mode Nuit :
        ' On vérifie dynamiquement SEULEMENT le personnel de Nuit (31-38) et les remplacements Nuit (46-47).
        dynamicRows = Array(31, 32, 33, 34, 35, 36, 37, 38, 46, 47)
        
        ' On masque de FORCE tous les blocs non désirés, COMME DEMANDÉ.
        rowsToHide = Array("5:28", "39:45", "48:58", "60:62", "64:70")
    End If

    ' 3. EXÉCUTION DES ACTIONS DE MASQUAGE
    HideRowBlocks ws, rowsToHide
    HandleDynamicRows ws, dynamicRows
    AutoHideRowsBasedOnName ws ' Gardé comme sécurité supplémentaire

    ' 4. FINALISATION DE L'AFFICHAGE
    
    ' On masque les commentaires qui se trouvent sur les lignes maintenant cachées.
    HideCommentsInHiddenRows ws

    ' --- NOUVELLE LIGNE POUR MASQUER LA COLONNE B ---
    ws.Columns("B").Hidden = True
    ' ------------------------------------------------
    
    ws.Columns(MENU_COLS).Hidden = True
    ActiveWindow.Zoom = 70
    
    ' Positionne la vue au bon endroit
    If mode = ViewJour Then
        Application.Goto ws.Range("A1"), Scroll:=True
    Else
        Application.Goto ws.Range("A30"), Scroll:=True ' Centre sur le personnel de nuit
    End If
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur est survenue dans la macro '" & IIf(mode = ViewJour, "Mode Jour", "Mode Nuit") & "'." _
           & vbCrLf & vbCrLf & "Erreur " & Err.Number & ": " & Err.Description, vbCritical, "Erreur d'exécution"
    Resume Cleanup
End Sub

' =====================================================================================
'   SOUS-ROUTINES PRIVÉES (OUTILS INTERNES)
' =====================================================================================
Private Sub AutoHideRowsBasedOnName(ws As Worksheet)
    Dim cell As Range
    On Error Resume Next
    For Each cell In ws.Range(NAME_RANGE_TO_CHECK)
        If cell.EntireRow.Hidden = False Then
            If Len(Trim(CStr(cell.value))) = 0 Then cell.EntireRow.Hidden = True
        End If
    Next cell
    On Error GoTo 0
End Sub

Private Sub HideRowBlocks(ws As Worksheet, blockArray As Variant)
    Dim blk As Variant
    On Error Resume Next
    For Each blk In blockArray
        ws.Rows(blk).Hidden = True
    Next blk
    On Error GoTo 0
End Sub

Private Sub HandleDynamicRows(ws As Worksheet, rowsArray As Variant)
    Dim r As Variant
    Dim nameCell As Range, scheduleCells As Range
    Dim isNameEmpty As Boolean, areSchedulesEmpty As Boolean
    
    On Error Resume Next
    For Each r In rowsArray
        If ws.Rows(r).Hidden = False Then ' Ne traite que les lignes encore visibles
            Set nameCell = ws.Cells(r, "A")
            Set scheduleCells = Intersect(ws.Rows(r), ws.Columns(START_SCHEDULE_COL & ":" & END_SCHEDULE_COL))
            isNameEmpty = (Len(Trim(CStr(nameCell.value))) = 0)
            areSchedulesEmpty = (Application.WorksheetFunction.CountA(scheduleCells) = 0)
            ws.Rows(r).Hidden = (isNameEmpty And areSchedulesEmpty)
        End If
    Next r
    On Error GoTo 0
End Sub
' =====================================================================================
'   SOUS-ROUTINE POUR MASQUER LES COMMENTAIRES DANS LES LIGNES CACHÉES
' =====================================================================================
Private Sub HideCommentsInHiddenRows(ws As Worksheet)
    Dim cmt As Comment
    
    On Error Resume Next ' Au cas où il n'y a aucun commentaire
    
    ' Parcourt chaque commentaire sur la feuille active
    For Each cmt In ws.Comments
        ' Vérifie si la ligne du commentaire est masquée
        If cmt.Parent.EntireRow.Hidden Then
            ' Si la ligne est masquée, on masque aussi la "forme" du commentaire
            cmt.Shape.Visible = False
        End If
    Next cmt
    
    On Error GoTo 0
End Sub


