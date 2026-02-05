Attribute VB_Name = "Debug_Noms"
'Attribute VB_Name = "Debug_Noms"
Option Explicit

Sub Debug_Verifier_Noms()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim wsPersonnel As Worksheet
    
    On Error Resume Next
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If wsPersonnel Is Nothing Then
        MsgBox "Feuille Personnel introuvable!", vbCritical
        Exit Sub
    End If
    
    ' Charger dictionnaire depuis Personnel
    Dim dictFct As Object: Set dictFct = CreateObject("Scripting.Dictionary")
    dictFct.CompareMode = vbTextCompare
    
    Dim lr As Long: lr = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
    Dim arr As Variant: arr = wsPersonnel.Range("B2:E" & lr).value
    Dim i As Long
    
    Dim rapport As String
    rapport = "=== PERSONNEL ===" & vbCrLf
    
    For i = 1 To UBound(arr, 1)
        Dim nom As String: nom = Trim(CStr(arr(i, 1)))
        Dim prenom As String: prenom = Trim(CStr(arr(i, 2)))
        Dim fct As String: fct = Trim(CStr(arr(i, 4)))
        Dim cle As String: cle = nom & "_" & prenom
        
        rapport = rapport & "Cle=[" & cle & "] Fct=[" & fct & "]" & vbCrLf
        
        If cle <> "_" And Not dictFct.Exists(cle) Then
            dictFct.Add cle, UCase(fct)
        End If
    Next i
    
    rapport = rapport & vbCrLf & "=== PLANNING (Col A, lignes 6-28) ===" & vbCrLf
    
    For i = 6 To 28
        Dim nomPlanning As String: nomPlanning = Trim(CStr(ws.Cells(i, 1).value))
        Dim cleTest As String: cleTest = Replace(nomPlanning, " ", "_")
        Dim fctTrouvee As String
        
        If dictFct.Exists(cleTest) Then
            fctTrouvee = dictFct(cleTest)
        ElseIf dictFct.Exists(nomPlanning) Then
            fctTrouvee = dictFct(nomPlanning)
        Else
            fctTrouvee = "???NON TROUVE???"
        End If
        
        rapport = rapport & "L" & i & " Planning=[" & nomPlanning & "] => Cle=[" & cleTest & "] => Fct=[" & fctTrouvee & "]" & vbCrLf
    Next i
    
    Debug.Print rapport
    MsgBox "Rapport dans Fenetre Immediate (Ctrl+G)", vbInformation
End Sub

