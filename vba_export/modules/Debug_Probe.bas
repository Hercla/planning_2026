Attribute VB_Name = "Debug_Probe"
Sub ProbeCalculation()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim rName As Range
    Set rName = ws.Cells(6, 1) ' Hermann_Claude
    Dim rCode As Range
    Set rCode = ws.Cells(6, 3) ' Jeudi 1st (Col 3)
    
    Dim msg As String
    msg = "Name Cell: " & rName.Address & " = " & rName.value & vbCrLf
    msg = msg & "Code Cell: " & rCode.Address & " = " & rCode.value & " (Color: " & rCode.Interior.Color & ")" & vbCrLf
    
    ' Load Config
    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    Dim configGlobal As Object
    Set configGlobal = CreateObject("Scripting.Dictionary")
    If Not wsConfig Is Nothing Then
        ' Quick load simplified
        Dim arr, i
        arr = wsConfig.UsedRange.value
        For i = 1 To UBound(arr, 1)
            If Not IsError(arr(i, 1)) And Not IsError(arr(i, 2)) Then
                configGlobal(Trim(CStr(arr(i, 1)))) = arr(i, 2)
            End If
        Next i
    End If
    
    Dim fctCfg As String
    fctCfg = ""
    If configGlobal.Exists("CHK_InfFunctions") Then fctCfg = configGlobal("CHK_InfFunctions")
    msg = msg & "Config 'CHK_InfFunctions': " & fctCfg & vbCrLf
    
    ' Load Personnel
    Dim dictFonctions As Object
    Set dictFonctions = CreateObject("Scripting.Dictionary")
    Dim wsPerso As Worksheet
    On Error Resume Next
    Set wsPerso = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If Not wsPerso Is Nothing Then
        Dim lastRow As Long
        lastRow = wsPerso.Cells(wsPerso.Rows.count, "B").End(xlUp).row
        msg = msg & "Personnel Sheet found, rows: " & lastRow & vbCrLf
        Dim arrP
        arrP = wsPerso.Range("B2:F" & lastRow).value
        Dim nom As String, prenom As String, fct As String
        
        ' Manual check for Hermann
        Dim found As Boolean
        For i = 1 To UBound(arrP, 1)
            nom = Trim(CStr(arrP(i, 1)))
            prenom = Trim(CStr(arrP(i, 2)))
            fct = Trim(CStr(arrP(i, 5)))
            dictFonctions(nom & "_" & prenom) = fct
            
            If InStr(rName.value, nom) > 0 Then
                msg = msg & "   -> Match found in Personnel: " & nom & " " & prenom & " = " & fct & vbCrLf
                found = True
            End If
        Next i
        If Not found Then msg = msg & "   -> NO Match found in Personnel for " & rName.value & vbCrLf
    Else
        msg = msg & "Personnel Sheet NOT FOUND!" & vbCrLf
    End If
    
    ' Check Logic
    Dim nomP As String
    nomP = Trim(CStr(rName.value))
    Dim fPersonne As String
    fPersonne = ""
    If dictFonctions.Exists(Replace(nomP, " ", "_")) Then
        fPersonne = UCase(dictFonctions(Replace(nomP, " ", "_")))
    ElseIf dictFonctions.Exists(nomP) Then
        fPersonne = UCase(dictFonctions(nomP))
    End If
    
    msg = msg & "Resolved Function: '" & fPersonne & "'" & vbCrLf
    
    ' Check Exclusions
    Dim compter As Boolean: compter = False
    Dim fList As String
    If fctCfg = "" Then fList = "INF,AS,CEFA" Else fList = fctCfg
    fList = "," & Replace(UCase(fList), " ", "") & ","
    fList = Replace(fList, ";", ",")
    
    If InStr(fList, "," & UCase(fPersonne) & ",") > 0 Then
        compter = True
    End If
    msg = msg & "In List " & fList & " ? " & compter & vbCrLf
    
    MsgBox msg
End Sub
