Attribute VB_Name = "ReimportViewFixes"
Option Explicit

' ============================================================================
' SCRIPT DE REIMPORT - FIX MODE JOUR/NUIT + COLONNE B
' Importe les modules corriges depuis vba_export
' ============================================================================

Sub ReimportViewModules()
    Dim vbp As Object
    Dim modulePath As String
    Dim basePath As String

    ' Chemin vers le dossier vba_export
    basePath = "C:\Users\hercl\planning_2026_repo\vba_export\"

    Set vbp = ThisWorkbook.VBProject

    On Error Resume Next

    ' 1. Supprimer les anciens modules
    vbp.VBComponents.Remove vbp.VBComponents("Module_Calculer_Totaux")
    vbp.VBComponents.Remove vbp.VBComponents("MODULEMODES_CONFIGDRIVEN")

    On Error GoTo 0

    ' 2. Reimporter les modules corriges
    On Error Resume Next

    vbp.VBComponents.Import basePath & "Module_Calculer_Totaux.bas"
    If Err.Number <> 0 Then
        MsgBox "Erreur import Module_Calculer_Totaux: " & Err.Description, vbExclamation
        Err.Clear
    End If

    vbp.VBComponents.Import basePath & "MODULEMODES_CONFIGDRIVEN.bas"
    If Err.Number <> 0 Then
        MsgBox "Erreur import MODULEMODES_CONFIGDRIVEN: " & Err.Description, vbExclamation
        Err.Clear
    End If

    On Error GoTo 0

    MsgBox "Reimport termine !" & vbCrLf & vbCrLf & _
           "Modules mis a jour :" & vbCrLf & _
           "  - Module_Calculer_Totaux (fix restauration mode)" & vbCrLf & _
           "  - MODULEMODES_CONFIGDRIVEN (colonne B masquee en Mode J)", vbInformation
End Sub
