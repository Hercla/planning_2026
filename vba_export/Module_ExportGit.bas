Attribute VB_Name = "Module_ExportGit"
Option Explicit

' ===================================================================
' EXPORT AUTOMATIQUE VBA ? REPO GIT
' Usage : ExportAllVBAToRepo après chaque modif validée
' ===================================================================

Public Sub ExportAllVBAToRepo()
    Const REPO_PATH As String = "C:\Users\hercl\planning_2026_repo\vba_export\"
    
    Dim vbComp As Object ' VBComponent
    Dim exportPath As String
    Dim extension As String
    Dim exportCount As Long
    
    ' Vérifier que le dossier repo existe
    If Dir(REPO_PATH, vbDirectory) = "" Then
        MsgBox "ERREUR : Dossier repo introuvable" & vbCrLf & REPO_PATH, vbCritical
        Exit Sub
    End If
    
    ' Nettoyer le dossier export (optionnel - décommenter si besoin)
    ' CleanExportFolder REPO_PATH
    
    ' Exporter chaque composant
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule
                extension = ".bas"
            Case 2 ' vbext_ct_ClassModule
                extension = ".cls"
            Case 3 ' vbext_ct_MSForm
                extension = ".frm"
            Case 100 ' vbext_ct_Document (feuilles/ThisWorkbook)
                extension = ".cls"
            Case Else
                extension = "" ' Skip
        End Select
        
        If extension <> "" Then
            exportPath = REPO_PATH & vbComp.Name & extension
            
            ' Export avec écrasement
            On Error Resume Next
            vbComp.Export exportPath
            
            If Err.Number = 0 Then
                exportCount = exportCount + 1
                Debug.Print "? " & vbComp.Name & extension
            Else
                Debug.Print "? ERREUR : " & vbComp.Name & " - " & Err.description
            End If
            On Error GoTo 0
        End If
    Next vbComp
    
    MsgBox exportCount & " modules exportés vers repo" & vbCrLf & REPO_PATH, vbInformation, "Export Git"
End Sub

' Optionnel : Nettoyer dossier avant export (évite fichiers fantômes)
Private Sub CleanExportFolder(folderPath As String)
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each file In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(file.Name)) Like "bas" Or _
           LCase(fso.GetExtensionName(file.Name)) Like "cls" Or _
           LCase(fso.GetExtensionName(file.Name)) Like "frm" Then
            file.Delete
        End If
    Next file
End Sub

