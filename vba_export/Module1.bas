Attribute VB_Name = "Module1"
Sub P0_FixSplitQuotes()
    Debug.Print "=== FIX SPLIT QUOTES ==="
    
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents("Module_ParseHoraire").CodeModule
    
    Dim i As Long, line As String
    Dim fixCount As Long
    
    For i = 1 To cm.CountOfLines
        line = cm.lines(i, 1)
        
        ' Chercher Split avec problème guillemets
        If InStr(line, "Split(defStr,") > 0 Then
            ' Remplacer par version correcte
            line = Replace(line, "Split(defStr, "" "")", "Split(defStr, "" "")")
            line = Replace(line, "Split(defStr, """")", "Split(defStr, "" "")")
            
            ' Version safe : forcer syntaxe correcte
            If InStr(line, "parts = Split") > 0 Then
                cm.ReplaceLine i, "    parts = Split(defStr, "" "")"
                Debug.Print "? L" & i & " : Split corrigé"
                fixCount = fixCount + 1
            End If
        End If
    Next i
    
    Debug.Print "=== " & fixCount & " CORRECTIONS ==="
    Debug.Print "RECOMPILER MAINTENANT"
End Sub
