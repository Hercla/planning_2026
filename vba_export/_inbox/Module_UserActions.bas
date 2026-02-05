' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module_UserActions"
Option Explicit

'================================================================================================
' MODULE :          Module_UserActions
' DESCRIPTION :     Centralise TOUTES les actions déclenchées par l'interface utilisateur
'                   (boutons sur la feuille, boutons sur le UserForm, cases à cocher).
'================================================================================================


' --- CONSTANTES PRIVÉES ---
Private Const PLANNING_RANGE_NAME As String = "planning"


'================================================================================================
'   SECTION 1: ACTIONS POUR LES BOUTONS SUR LA FEUILLE EXCEL
'================================================================================================

Public Sub InsertCodeFromSheetButton()
    ' DESCRIPTION: Procédure unique affectée aux boutons (Formes) sur une feuille.
    '              Elle lit le texte du bouton cliqué pour déterminer quel code insérer.
    
    If Intersect(ActiveCell, Range(PLANNING_RANGE_NAME)) Is Nothing Then Exit Sub
    
    Dim codeToInsert As String
    On Error Resume Next
    codeToInsert = ActiveSheet.Shapes(Application.Caller).TextFrame.Characters.text
    On Error GoTo 0
    
    If codeToInsert = "" Then Exit Sub
    
    ApplyCodeAndMove codeToInsert
End Sub

Public Sub ToggleReplacementRows()
    ' DESCRIPTION: Affiche ou masque les lignes en fonction de la case à cocher.
    Dim cb As CheckBox
    On Error Resume Next
    Set cb = ActiveSheet.CheckBoxes(Application.Caller)
    On Error GoTo 0
    
    If cb Is Nothing Then Exit Sub
    ActiveSheet.Rows("43:44").EntireRow.Hidden = (cb.value <> xlOn)
End Sub


'================================================================================================
'   SECTION 2: ACTIONS POUR LES BOUTONS SUR LE USERFORM
'================================================================================================

Public Sub InsertCodeFromUserForm(ByVal code As String)
    ' DESCRIPTION: Procédure appelée par les boutons du UserForm.
    '              Elle reçoit le code à insérer en paramètre.
    
    If Intersect(ActiveCell, Range(PLANNING_RANGE_NAME)) Is Nothing Then Exit Sub
    ApplyCodeAndMove code
End Sub


'================================================================================================
'   SECTION 3: ACTIONS DE NAVIGATION
'================================================================================================

Public Sub NavigateToSheet(ByVal sheetName As String)
    ' DESCRIPTION: Procédure unique pour activer une feuille et se positionner.
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La feuille '" & sheetName & "' est introuvable.", vbExclamation
        Exit Sub
    End If
    
    ws.Activate
    ws.Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub


'================================================================================================
'   PROCÉDURE D'AIDE PRIVÉE (INTERNE AU MODULE)
'================================================================================================

Private Sub ApplyCodeAndMove(ByVal code As String)
    ' DESCRIPTION: Procédure centrale qui insère le code, formate la cellule et se déplace.
    
    With ActiveCell
        .value = code
        .Interior.Color = vbWhite
        .Font.Color = vbBlack
        On Error Resume Next
        .offset(0, 1).Select
        On Error GoTo 0
    End With
End Sub
