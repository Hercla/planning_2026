Attribute VB_Name = "Debug_PM_Detail"
'Attribute VB_Name = "Debug_PM_Detail"
Option Explicit

Sub Debug_PM_Jour31()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim col As Long: col = 33  ' AG = colonne 33
    
    Dim wsConfig As Worksheet, wsPersonnel As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    ' Config
    Dim ligneDebut As Long: ligneDebut = 6
    Dim ligneFin As Long: ligneFin = 28
    Dim couleurIgnore As Long: couleurIgnore = 15849925
    Dim couleurJaune As Long: couleurJaune = 65535
    Dim couleurBleu As Long: couleurBleu = 15128749
    
    ' Charger fonctions
    Dim dictFct As Object: Set dictFct = CreateObject("Scripting.Dictionary")
    dictFct.CompareMode = vbTextCompare
    If Not wsPersonnel Is Nothing Then
        Dim lr As Long: lr = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
        Dim arr As Variant: arr = wsPersonnel.Range("B2:E" & lr).value
        Dim i As Long
        For i = 1 To UBound(arr, 1)
            Dim cle As String: cle = Trim(CStr(arr(i, 1))) & "_" & Trim(CStr(arr(i, 2)))
            If cle <> "_" And Not dictFct.Exists(cle) Then
                dictFct.Add cle, UCase(Trim(CStr(arr(i, 4))))
            End If
        Next i
    End If
    
    Dim rapport As String
    rapport = "=== DEBUG PM - JOUR 31 (Col AG) - " & ws.Name & " ===" & vbCrLf & vbCrLf
    
    Dim totalPM As Double: totalPM = 0
    Dim countDetails As String: countDetails = ""
    
    For i = 6 To 28
        Dim nomPersonne As String: nomPersonne = Trim(CStr(ws.Cells(i, 1).value))
        Dim code As String: code = Trim(CStr(ws.Cells(i, col).value))
        Dim couleur As Long: couleur = ws.Cells(i, col).Interior.Color
        
        If Len(code) = 0 Then GoTo NextRow
        
        ' Fonction
        Dim fct As String: fct = ""
        If dictFct.Exists(nomPersonne) Then fct = dictFct(nomPersonne)
        
        ' Exclusions
        Dim exclu As Boolean: exclu = False
        Dim raison As String: raison = ""
        
        If couleur = couleurIgnore Then exclu = True: raison = "ROUGE"
        If couleur = couleurJaune Then exclu = True: raison = "JAUNE"
        If couleur = couleurBleu Then exclu = True: raison = "BLEU"
        
        Dim codeUp As String: codeUp = UCase(code)
        If codeUp = "WE" Or Left(codeUp, 3) = "MAL" Or Left(codeUp, 2) = "CA" Or _
           Left(codeUp, 3) = "RCT" Or codeUp = "DP" Or codeUp = "RHS" Then
            exclu = True: raison = "ABSENCE"
        End If
        
        If fct <> "INF" And fct <> "AS" And fct <> "CEFA" Then
            exclu = True: raison = "FCT=" & fct
        End If
        
        ' Calculer PM (fin > 13h)
        Dim worksAM As Double: worksAM = 0
        If Not exclu Then
            ' Parser le code pour trouver heure de fin
            Dim parts() As String
            Dim finHeure As Double: finHeure = 0
            
            code = Replace(code, vbLf, " ")
            code = Replace(code, vbCr, " ")
            Do While InStr(code, "  ") > 0: code = Replace(code, "  ", " "): Loop
            parts = Split(Trim(code), " ")
            
            ' Trouver la derniere heure (fin)
            Dim p As Long
            For p = UBound(parts) To 0 Step -1
                If InStr(parts(p), ":") > 0 Or IsNumeric(parts(p)) Then
                    finHeure = HeureEnDecimal(parts(p))
                    If finHeure > 0 Then Exit For
                End If
            Next p
            
            ' PM = fin > 13h
            If finHeure > 13 Then worksAM = 1
            
            rapport = rapport & "L" & i & " " & Left(nomPersonne & String(25, " "), 25)
            rapport = rapport & " | " & Left(code & String(15, " "), 15)
            rapport = rapport & " | Fin=" & Format(finHeure, "0.00")
            rapport = rapport & " | PM=" & worksAM
            rapport = rapport & " | " & fct & vbCrLf
            
            totalPM = totalPM + worksAM
            If worksAM > 0 Then countDetails = countDetails & nomPersonne & ", "
        Else
            rapport = rapport & "L" & i & " " & Left(nomPersonne & String(25, " "), 25)
            rapport = rapport & " | " & Left(code & String(15, " "), 15)
            rapport = rapport & " | EXCLU: " & raison & vbCrLf
        End If
        
NextRow:
    Next i
    
    rapport = rapport & vbCrLf & String(60, "-") & vbCrLf
    rapport = rapport & "TOTAL PM CALCULE: " & totalPM & vbCrLf
    rapport = rapport & "Qui compte: " & countDetails & vbCrLf
    
    Debug.Print rapport
    MsgBox "Total PM = " & totalPM & vbCrLf & vbCrLf & "Details dans Fenetre Immediate (Ctrl+G)", vbInformation
End Sub

Private Function HeureEnDecimal(s As String) As Double
    HeureEnDecimal = 0
    If Len(s) = 0 Then Exit Function
    
    If InStr(s, ":") > 0 Then
        Dim parts() As String: parts = Split(s, ":")
        On Error Resume Next
        HeureEnDecimal = CDbl(parts(0)) + CDbl(parts(1)) / 60
        On Error GoTo 0
    ElseIf IsNumeric(s) Then
        Dim v As Double: v = CDbl(s)
        If v < 1 And v > 0 Then
            HeureEnDecimal = v * 24
        Else
            HeureEnDecimal = v
        End If
    End If
End Function


