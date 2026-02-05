' Debug pour vérifier si ParseCode fonctionne
Sub Debug_ParseCode()
    Dim codes As Variant
    codes = Array("8:30 16:30", "7 15:30", "7 13", "6:45 15:15", "C 20 E")
    
    Dim code As Variant
    Dim h1 As Double, f1 As Double, h2 As Double, f2 As Double
    Dim mat As Double, am As Double, soi As Double, nui As Double
    Dim msg As String
    
    For Each code In codes
        If ParseCodeDebug(CStr(code), h1, f1, h2, f2) Then
            CalcPeriodesDebug h1, f1, h2, f2, mat, am, soi, nui
            msg = msg & code & ":" & vbLf
            msg = msg & "  Heures: " & h1 & "-" & f1 & " / " & h2 & "-" & f2 & vbLf
            msg = msg & "  Matin=" & mat & ", AM=" & am & ", Soir=" & soi & ", Nuit=" & nui & vbLf & vbLf
        Else
            msg = msg & code & ": PARSE FAILED" & vbLf & vbLf
        End If
    Next code
    
    MsgBox msg, vbInformation, "Debug ParseCode"
End Sub

Private Function ParseCodeDebug(c As String, ByRef s1 As Double, ByRef e1 As Double, ByRef s2 As Double, ByRef e2 As Double) As Boolean
    s1 = 0: e1 = 0: s2 = 0: e2 = 0
    Dim p() As String, tmp As String
    tmp = Trim(Replace(Replace(c, vbLf, " "), vbCr, " "))
    Do While InStr(tmp, "  ") > 0: tmp = Replace(tmp, "  ", " "): Loop
    p = Split(tmp, " ")
    
    On Error GoTo Err1
    If UBound(p) = 1 Then
        s1 = HeureDecDebug(p(0)): e1 = HeureDecDebug(p(1))
        ParseCodeDebug = True
    ElseIf UBound(p) >= 3 Then
        s1 = HeureDecDebug(p(0)): e1 = HeureDecDebug(p(1))
        s2 = HeureDecDebug(p(2)): e2 = HeureDecDebug(p(3))
        ParseCodeDebug = True
    Else
        ParseCodeDebug = False
    End If
    Exit Function
Err1:
    ParseCodeDebug = False
End Function

Private Function HeureDecDebug(s As String) As Double
    Dim p() As String
    If InStr(s, ":") > 0 Then
        p = Split(s, ":")
        HeureDecDebug = CDbl(p(0)) + CDbl(p(1)) / 60
    ElseIf IsNumeric(s) Then
        HeureDecDebug = CDbl(s)
    Else
        HeureDecDebug = 0
    End If
End Function

Private Sub CalcPeriodesDebug(h1 As Double, f1 As Double, h2 As Double, f2 As Double, ByRef mat As Double, ByRef am As Double, ByRef soi As Double, ByRef nui As Double)
    mat = 0: am = 0: soi = 0: nui = 0
    Dim heureMatin As Double, heureAM As Double
    
    ' Matin = 8h-12h (4 heures max)
    heureMatin = OverlapDebug(h1, f1, 8, 12) + OverlapDebug(h2, f2, 8, 12)
    If heureMatin >= 4 Then
        mat = 1
    ElseIf heureMatin >= 2 Then
        mat = 0.5
    ElseIf heureMatin > 0 Then
        mat = Round(heureMatin / 4, 2)
    End If
    
    ' AM = 12h-16h30 (4.5 heures max)
    heureAM = OverlapDebug(h1, f1, 12, 16.5) + OverlapDebug(h2, f2, 12, 16.5)
    If heureAM >= 4 Then
        am = 1
    ElseIf heureAM >= 2 Then
        am = 0.5
    ElseIf heureAM > 0 Then
        am = Round(heureAM / 4.5, 2)
    End If
    
    ' Soir = après 16h30
    If f1 > 16.5 Or f2 > 16.5 Then soi = 1
    
    ' Nuit = commence après 19h30 OU finit avant 7h15
    If h1 >= 19.5 Or f1 <= 7.25 Then nui = 1
End Sub

Private Function OverlapDebug(hd As Double, hf As Double, td As Double, tf As Double) As Double
    Dim os As Double, oe As Double
    os = Application.Max(hd, td): oe = Application.Min(hf, tf)
    If oe > os Then OverlapDebug = oe - os Else OverlapDebug = 0
End Function
