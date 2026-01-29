Attribute VB_Name = "Module_EmergencyExport"
Option Explicit

Public Sub EmergencyExport()
    Const REPO As String = "C:\Users\hercl\planning_2026_repo\vba_export_dirty\"
    Dim vbComp As Object, ext As String, p As String
    
    On Error Resume Next
    MkDir REPO
    On Error GoTo 0
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: ext = ".bas"         ' standard module
            Case 2, 100: ext = ".cls"    ' class / document
            Case 3: ext = ".frm"         ' userform
            Case Else: ext = ""
        End Select
        
        If ext <> "" Then
            p = REPO & vbComp.Name & ext
            On Error Resume Next
            vbComp.Export p
            Debug.Print IIf(Err.Number = 0, "? ", "? ") & vbComp.Name & ext
            Err.Clear
            On Error GoTo 0
        End If
    Next vbComp
    
    MsgBox "Export DIRTY terminé : " & vbCrLf & REPO, vbInformation
End Sub

