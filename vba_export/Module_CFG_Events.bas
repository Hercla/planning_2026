Attribute VB_Name = "Module_CFG_Events"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

'===================================================================================
' MODULE: Module_CFG_Events (BEST - Local-First)
' PURPOSE:
'   Debounce + garde-fou pour appliquer la vue APRÈS changements de CFG,
'   mais UNIQUEMENT sur l’onglet actif (local-first).
'===================================================================================

Private gInCfgChange As Boolean
Private gPendingViewApply As Boolean
Private gNextRunAt As Date

' Appelé par Feuil_Config.Worksheet_Change
Public Sub CFG_OnChange_RequestViewApply()
    If gInCfgChange Then Exit Sub
    
    gPendingViewApply = True
    gNextRunAt = Now + TimeSerial(0, 0, 1) ' debounce 1s
    
    On Error Resume Next
    Application.OnTime EarliestTime:=gNextRunAt, _
                       Procedure:="CFG_ApplyView_IfPending", _
                       schedule:=True
    On Error GoTo 0
End Sub

' Exécution réelle (debounced)
Public Sub CFG_ApplyView_IfPending()
    If gInCfgChange Then Exit Sub
    If Not gPendingViewApply Then Exit Sub
    
    gInCfgChange = True
    On Error GoTo CleanUp
    
    gPendingViewApply = False
    ' LOCAL-FIRST : on n’applique QUE sur l’onglet actif
    VIEW_Apply_ByScope
    
CleanUp:
    gInCfgChange = False
End Sub

' À appeler à la fermeture pour annuler OnTime éventuel
Public Sub CFG_OnWorkbookClose_CancelPending()
    On Error Resume Next
    If gNextRunAt <> 0 Then
        Application.OnTime EarliestTime:=gNextRunAt, _
                           Procedure:="CFG_ApplyView_IfPending", _
                           schedule:=False
    End If
    gPendingViewApply = False
    On Error GoTo 0
End Sub


