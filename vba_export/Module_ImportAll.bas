Attribute VB_Name = "Module_ImportAll"
Option Explicit

'=========================================================================
' ImportAllVBA
' - Importe tous les .bas .cls .frm depuis un dossier
' - Supprime/replace les composants existants (sauf ThisWorkbook/Sheets)
'
' Prérequis Excel:
'   Fichier > Options > Centre de gestion de la confidentialité >
'   Paramètres du centre > Paramètres des macros >
'   ? "Accès approuvé au modèle d'objet du projet VBA"
'=========================================================================

Public Sub ImportAllVBA()
    Dim folderPath As String
    folderPath = "C:\Users\hercl\planning_2026"   ' <-- adapte si besoin
    
    Dim fso As Object, folder As Object, fil As Object
    Dim proj As Object, vbComp As Object
    
    Dim ext As String
    Dim compName As String
    
    On Error GoTo EH
    
    ' --- perf Excel
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Dossier introuvable : " & folderPath, vbCritical
        GoTo CleanExit
    End If
    Set folder = fso.GetFolder(folderPath)
    
    ' --- accès VBProject
    Set proj = ThisWorkbook.VBProject
    
    ' --- loop files
    For Each fil In folder.Files
        ext = LCase$(fso.GetExtensionName(fil.path))
        
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            
            compName = fso.GetBaseName(fil.path)
            
            ' Supprime composant existant (si présent), sauf Documents (ThisWorkbook/Worksheets)
            Set vbComp = Nothing
            On Error Resume Next
            Set vbComp = proj.VBComponents(compName)
            On Error GoTo EH
            
            If Not vbComp Is Nothing Then
                ' vbext_ct_Document = 100 (ThisWorkbook/Sheets) ? ne pas supprimer
                If vbComp.Type <> 100 Then
                    proj.VBComponents.Remove vbComp
                End If
            End If
            
            ' Import
            proj.VBComponents.Import fil.path
        End If
    Next fil
    
    MsgBox "Import terminé ?" & vbCrLf & "Source : " & folderPath, vbInformation

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

EH:
    ' Erreur typique si "Trust access..." désactivé : 1004 / permission
    MsgBox "Erreur ImportAllVBA ?" & vbCrLf & _
           "Détail: " & Err.Number & " - " & Err.description & vbCrLf & vbCrLf & _
           "Vérifie: Options > Centre de gestion de la confidentialité > Paramètres des macros > " & _
           """Accès approuvé au modèle d'objet du projet VBA""", vbCritical
    Resume CleanExit
End Sub


