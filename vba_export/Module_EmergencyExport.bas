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
Sub DetailedScan()
    Dim vbComp As Object, cm As Object
    Dim i As Long, line As String
    Dim issueCount As Long, totalLines As Long
    
    Debug.Print "=== SCAN DÉTAILLÉ 142 MODULES ==="
    Debug.Print String(60, "=")
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Set cm = vbComp.CodeModule
        issueCount = 0
        totalLines = cm.CountOfLines
        
        ' Check Option Explicit
        If totalLines > 0 Then
            line = cm.lines(1, 1)
            If InStr(1, line, "Option Explicit", vbTextCompare) = 0 Then
                Debug.Print "?? " & vbComp.Name & " - Manque Option Explicit"
                issueCount = issueCount + 1
            End If
        End If
        
        ' Scan patterns problématiques
        For i = 1 To totalLines
            line = cm.lines(i, 1)
            
            ' Références Feuil1 (à migrer vers Feuil_Config)
            If InStr(line, "Feuil1.") > 0 Or InStr(line, "Sheets(""Feuil1"")") > 0 Then
                Debug.Print "  L" & i & " ? Ref Feuil1: " & Left(Trim(line), 50)
                issueCount = issueCount + 1
            End If
            
            ' #REF hardcodés
            If InStr(line, "#REF") > 0 Then
                Debug.Print "  L" & i & " ? #REF: " & Left(Trim(line), 50)
                issueCount = issueCount + 1
            End If
        Next i
        
        If issueCount > 5 Then
            Debug.Print "  ?? " & vbComp.Name & " - " & issueCount & " problèmes détectés"
        ElseIf issueCount > 0 Then
            Debug.Print "  ?? " & vbComp.Name & " - " & issueCount & " problèmes"
        End If
    Next vbComp
    
    Debug.Print String(60, "=")
    Debug.Print "? Scan terminé - Scroll up pour voir détails"
End Sub
