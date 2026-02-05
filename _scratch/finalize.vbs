Option Explicit
Dim xlApp, xlBk, selectedFile
Dim pathModuleFix
Dim fso

pathModuleFix = "c:\Users\hercl\planning_2026\Module_Fix_Style.bas"

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(pathModuleFix) Then
    MsgBox "Erreur: Le fichier Module_Fix_Style.bas est introuvable.", vbCritical
    WScript.Quit
End If

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True

MsgBox "Veuillez selectionner le fichier Planning pour finaliser la mise en forme.", vbInformation

With xlApp.FileDialog(3)
    .Filters.Clear: .Filters.Add "Excel Files", "*.xlsm"
    If .Show = -1 Then selectedFile = .SelectedItems(1) Else WScript.Quit
End With

Set xlBk = xlApp.Workbooks.Open(selectedFile)

' Import Fix Module
On Error Resume Next
xlBk.VBProject.VBComponents.Import pathModuleFix
On Error GoTo 0

' Run Fix
xlApp.Run "Finaliser_Migration_Style"

' Save
xlBk.Save
MsgBox "Terminé ! Le style est corrigé et le calcul a été lancé.", vbInformation
