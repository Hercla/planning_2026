Attribute VB_Name = "FillSchedule"
' ExportedAt: 2026-01-12 15:37:08 | Workbook: Planning_2026.xlsm
Sub FillSchedule()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim timeRanges As Object
    Dim zeroCodes As Object
    Dim startTime As Double, endTime As Double
    Dim colIndex As Integer
    
    Set ws = ThisWorkbook.Sheets("Liste")
    Set rng = ws.Range("A2:A" & ws.Cells(ws.Rows.count, "A").End(xlUp).row)

    ' Define time ranges
    Set timeRanges = CreateObject("Scripting.Dictionary")
    timeRanges.Add "Matin", Array(8, 12)
    timeRanges.Add "Après-midi", Array(12, 16)
    timeRanges.Add "Soir", Array(16, 20)
    timeRanges.Add "Nuit", Array(20, 8)

    ' Define codes that are always 0
    Set zeroCodes = CreateObject("Scripting.Dictionary")
    Dim codeToAdd As Variant
    For Each codeToAdd In Array("FP", "CEP", "CP", "1/2*", "3/4*", "4/5*", "WE", "AAIR", "AFC", _
                                "ANC 1", "ANC 2", "ANC 3", "ANC 4", "ANC 5", "ANC 6", "ANC 7", "ANC 8", _
                                "BUS 1", "BUS 2", "BUS 3", "BUS 4", "CL1", "CL2", "CL3", "CL4", "CL5", "CL6", _
                                "CL7", "CL8", "CL9", "CL10", "CL11", "CL12", "CL13", "CL14", "CL15", "CL16", _
                                "CL17", "CL18", "CL19", "CL20", "CTR 1", "CTR 2", "CTR 3", "CTR 4", "CTR 5", _
                                "CTR 6", "CTR 7", "CTR 8", "CTR 9", "CTR 10", "CTR 11", "CTR 12", "EL1", "EL2", _
                                "EL3", "EL4", "EL5", "CS 4,57", "CS 5,4", "CS 6h", "CS 7,60", "F 3h30", "F 4h", _
                                "F 7h30", "F 6h30", "F 5h30", "F 7h", "F 8h", "FSH", "M 3h48", "M 6h", "M 7h36", _
                                "M 7,6", "M 4h", "M 5h42", "M 7h", "M 8h", "M 11h", "PETIT CHOM", "C ss solde 24min", _
                                "C ss solde 2h", "C ss solde 4h", "C ss solde 6h", "C ss solde 7,6", "Décès", _
                                "EM 3,8", "EM 6h", "EM 7,6", "Pat 6h", "Pat 7,6", "Préavis 3h48", "Préavis 6h", _
                                "Préavis 7h36", "VJ 7,6", "R.AFC", "RCT", "RHS 2h", "RHS 3h", "RHS 4h", "RHS 5h", _
                                "RHS 6h", "RHS 8h", "TV", "Déménag", "Grève")
        zeroCodes.Add codeToAdd, True
    Next codeToAdd

    For Each cell In rng
        Dim code As String
        code = Trim(cell.value) ' Trim spaces

        If Len(code) > 0 Then
            If zeroCodes.Exists(Left(code, 6)) Or zeroCodes.Exists(Left(code, 5)) _
               Or zeroCodes.Exists(Left(code, 4)) Or zeroCodes.Exists(Left(code, 3)) _
               Or zeroCodes.Exists(Left(code, 2)) Then
                ws.Cells(cell.row, 2).Resize(1, 4).value = 0
            Else
                ' Remove weekend abbreviations if present
                If InStr(code, "sa") > 0 Or InStr(code, "di") > 0 Then
                    code = Replace(Replace(code, "sa", ""), "di", "")
                End If
                
                Dim parts() As String
                parts = Split(code, " ")
                If UBound(parts) = 1 Then
                    Dim times() As String
                    times = Split(parts(1), ":")
                    If UBound(times) = 1 Then
                        If IsNumeric(times(0)) And IsNumeric(times(1)) Then
                            startTime = CDbl(times(0)) + CDbl(times(1)) / 60
                            endTime = startTime + 8
                            
                            ' Adjust end time if it wraps past midnight
                            If endTime >= 24 Then
                                endTime = endTime - 24
                            End If

                            For colIndex = 2 To 5
                                Dim rangeStart As Double, rangeEnd As Double
                                rangeStart = timeRanges.items()(colIndex - 2)(0)
                                rangeEnd = timeRanges.items()(colIndex - 2)(1)

                                If (rangeStart < rangeEnd And startTime >= rangeStart And endTime <= rangeEnd) Or _
                                   (rangeStart > rangeEnd And (startTime >= rangeStart Or endTime <= rangeEnd)) Then
                                    ws.Cells(cell.row, colIndex).value = 1
                                End If
                            Next colIndex
                        Else
                            Debug.Print "Error parsing time: " & times(0) & ":" & times(1)
                        End If
                    Else
                        Debug.Print "Unexpected time format: " & parts(1)
                    End If
                Else
                    Debug.Print "Unexpected code format: " & code
                End If
            End If
        End If
    Next cell
End Sub

