Attribute VB_Name = "ModPersonnelHelper"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

' Load unique full names from a Personnel sheet into a combo box.
' sheetName: name of the Personnel sheet
' colFirstName: column index containing first names
' colLastName: column index containing last names
' targetCombo: the combo box to populate
Public Sub LoadUniqueNames(ByVal sheetName As String, _
                           ByVal colFirstName As Long, _
                           ByVal colLastName As Long, _
                           ByRef targetCombo As MSForms.ComboBox)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Variant
    Dim namesDict As Object
    Dim i As Long
    Dim fullName As String
    Dim SortedKeys As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, colFirstName).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    dataRange = ws.Range(ws.Cells(2, colFirstName), ws.Cells(lastRow, colLastName)).value
    If Not IsArray(dataRange) Then Exit Sub

    Set namesDict = CreateObject("Scripting.Dictionary")
    namesDict.CompareMode = vbTextCompare

    For i = 1 To UBound(dataRange, 1)
        If Trim(dataRange(i, 1)) <> "" And Trim(dataRange(i, 2)) <> "" Then
            fullName = Trim(dataRange(i, 1)) & " " & Trim(dataRange(i, 2))
            If Not namesDict.Exists(fullName) Then namesDict.Add fullName, 1
        End If
    Next i

    If namesDict.count > 0 Then
        SortedKeys = SortDictionaryKeys(namesDict)
        targetCombo.list = SortedKeys
    End If
End Sub

' Return the sorted keys of a dictionary as a variant array
Private Function SortDictionaryKeys(ByVal dict As Object) As Variant
    If dict.count = 0 Then
        SortDictionaryKeys = Array()
        Exit Function
    End If

    Dim keysArray() As String
    Dim i As Long
    Dim key As Variant
    ReDim keysArray(0 To dict.count - 1)
    i = 0
    For Each key In dict.keys
        keysArray(i) = CStr(key)
        i = i + 1
    Next key
    QuickSort keysArray, LBound(keysArray), UBound(keysArray)
    SortDictionaryKeys = keysArray
End Function

' Simple QuickSort implementation for an array of strings
Private Sub QuickSort(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long
    Dim pivot As String, temp As String
    If first >= last Then Exit Sub
    low = first
    high = last
    pivot = arr((first + last) \ 2)
    Do While low <= high
        Do While StrComp(arr(low), pivot, vbTextCompare) < 0
            low = low + 1
        Loop
        Do While StrComp(arr(high), pivot, vbTextCompare) > 0
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub

