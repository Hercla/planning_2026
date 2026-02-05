Option Explicit
Dim xlApp, xlBk, fDialog, item, selectedFile
Dim pathModule, pathMigration
Dim fso

' Define paths to the BAS files we prepared
pathModule = "c:\Users\hercl\planning_2026\CalculFractionsPresence.bas"
pathMigration = "c:\Users\hercl\planning_2026\Module_Migration_Structure.bas"

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(pathModule) Or Not fso.FileExists(pathMigration) Then
    MsgBox "Erreur: Les fichiers .bas sources sont introuvables dans c:\Users\hercl\planning_2026\", vbCritical
    WScript.Quit
End If

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True ' Visible so user sees what happens

' Ask user to select file
MsgBox "Veuillez selectionner votre fichier Planning (ex: Planning_2026.xlsm) a mettre a jour.", vbInformation, "Selection Fichier"

' File Dialog implementation in VBS via Excel
Set fDialog = xlApp.FileDialog(3) ' msoFileDialogFilePicker
With fDialog
    .Title = "Selectionnez le fichier Excel Planning"
    .Filters.Clear
    .Filters.Add "Excel Macro-Enabled", "*.xlsm"
    .AllowMultiSelect = False
    If .Show = -1 Then
        selectedFile = .SelectedItems(1)
    Else
        MsgBox "Annule par l'utilisateur.", vbExclamation
        xlApp.Quit
        WScript.Quit
    End If
End With

On Error Resume Next
Set xlBk = xlApp.Workbooks.Open(selectedFile)
If Err.Number <> 0 Then
    MsgBox "Erreur a l'ouverture du fichier: " & Err.Description, vbCritical
    xlApp.Quit
    WScript.Quit
End If
On Error GoTo 0

' TRUST CHECK
Dim vbProj
On Error Resume Next
Set vbProj = xlBk.VBProject
If Err.Number <> 0 Then
    MsgBox "ERREUR CRITIQUE: L'acces au projet VBA est bloque." & vbCrLf & vbCrLf & _
           "SOLUTION: " & vbCrLf & _
           "1. Annulez ou fermez ce message." & vbCrLf & _
           "2. Dans Excel, allez dans Fichier > Options > Centre de gestion de la confidentialite > Parametres > Parametres des macros." & vbCrLf & _
           "3. COCHEZ 'Accès approuvé au modèle d'objet du projet VBA'." & vbCrLf & _
           "4. Relancez ce script.", vbCritical
    xlBk.Close False
    xlApp.Quit
    WScript.Quit
End If
On Error GoTo 0

' Remove old modules if they exist
Dim vbComp
On Error Resume Next
Set vbComp = xlBk.VBProject.VBComponents("CalculFractionsPresence")
If Not vbComp Is Nothing Then 
    xlBk.VBProject.VBComponents.Remove vbComp
    ' MsgBox "Ancien module CalculFractionsPresence supprime."
End If

Set vbComp = xlBk.VBProject.VBComponents("Module_Migration_Structure")
If Not vbComp Is Nothing Then 
    xlBk.VBProject.VBComponents.Remove vbComp
End If
On Error GoTo 0

' Import new modules
On Error Resume Next
xlBk.VBProject.VBComponents.Import pathModule
If Err.Number <> 0 Then MsgBox "Erreur import CalculFractionsPresence: " & Err.Description
xlBk.VBProject.VBComponents.Import pathMigration
If Err.Number <> 0 Then MsgBox "Erreur import Module_Migration_Structure: " & Err.Description
On Error GoTo 0

' Run Migration Logic
Dim reponse
reponse = MsgBox("Les modules de code ont ete mis a jour." & vbCrLf & vbCrLf & _
                 "Voulez-vous lancer la MIGRATION DE LA STRUCTURE maintenant ?" & vbCrLf & _
                 "(Cela va inserer les lignes 'Meteo' et 'INF' dans les onglets mois).", vbQuestion + vbYesNo, "Lancer Migration ?")

If reponse = vbYes Then
    On Error Resume Next
    xlApp.Run "Lancer_Migration_Totale"
    If Err.Number <> 0 Then
        MsgBox "Erreur durant la migration: " & Err.Description, vbCritical
    Else
        MsgBox "Migration de la structure terminee avec succes !", vbInformation
    End If
    On Error GoTo 0
Else
    MsgBox "Migration structurelle ignoree. Seul le code a ete mis a jour.", vbInformation
End If

' Save
xlBk.Save
MsgBox "Fichier sauvegarde. Le script va se fermer, mais Excel restera ouvert pour verification.", vbInformation

' Release objects
Set fDialog = Nothing
Set vbComp = Nothing
Set vbProj = Nothing
Set xlBk = Nothing
Set xlApp = Nothing
Set fso = Nothing
