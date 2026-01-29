Attribute VB_Name = "Module_UX_Mode"
' ExportedAt: 2026-01-12 16:03:00 | Workbook: Planning_2026.xlsm
Option Explicit

'===============================================================================
' MODULE: Module_UX_Mode
' PURPOSE:
'   UX MODE - Point d'entree pour boutons J/N
'   Appelle les macros Mode_Jour / Mode_Nuit de ModuleModes_ConfigDriven
'===============================================================================

'===============================================================================
' PUBLIC APIs - Fast Mode (pour boutons J/N)
'===============================================================================

Public Sub UX_FastModeJour()
    ' Appelle directement la macro Mode_Jour (config-driven)
    Mode_Jour
End Sub

Public Sub UX_FastModeNuit()
    ' Appelle directement la macro Mode_Nuit (config-driven)
    Mode_Nuit
End Sub

'===============================================================================
' Aliases supplementaires si besoin
'===============================================================================

Public Sub UX_Mode_Jour_ActiveSheet()
    Mode_Jour
End Sub

Public Sub UX_Mode_Nuit_ActiveSheet()
    Mode_Nuit
End Sub
