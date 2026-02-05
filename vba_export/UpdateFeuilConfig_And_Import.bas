Attribute VB_Name = "UpdateFeuilConfig_And_Import"
Option Explicit

  Public Sub UpdateFeuilConfig_And_Import()
      Const CSV_PATH As String = "C:\Users\hercl\planning-vba-automation\config\tblCFG.csv"
      Const MODULE_PATH As String = "C:\Users\hercl\planning-vba-automation\CalculFractionsPresence.bas"
      Const SHEET_NAME As String = "Feuil_Config"

      Dim ws As Worksheet
      Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

      Dim content As String
      content = ReadAllText(CSV_PATH)

      Dim lines() As String
      lines = Split(content, vbCrLf)

      Dim data() As Variant
      Dim rowCount As Long
      ReDim data(1 To UBound(lines), 1 To 2)

      Dim i As Long, line As String, fields As Variant
      For i = 1 To UBound(lines) ' skip header (index 0)
          line = Trim$(lines(i))
          If line = "" Then GoTo nextLine
          fields = ParseCsvLine(line)
          If UBound(fields) >= 1 Then
              rowCount = rowCount + 1
              data(rowCount, 1) = fields(0)
              data(rowCount, 2) = fields(1)
          End If
nextLine:
      Next i

      If rowCount = 0 Then
          MsgBox "CSV vide ou invalide.", vbExclamation
          Exit Sub
      End If

      ws.Range("A2:B" & ws.Rows.count).ClearContents
      ws.Range("A2").Resize(rowCount, 2).value = data

      ImportBasModule MODULE_PATH, "CalculFractionsPresence"

      MsgBox "Feuil_Config mise à jour + module importé.", vbInformation
  End Sub

  Private Sub ImportBasModule(ByVal modulePath As String, ByVal moduleName As String)
      Dim vbProj As Object, vbComp As Object
      Set vbProj = ThisWorkbook.VBProject

      On Error Resume Next
      Set vbComp = vbProj.VBComponents(moduleName)
      On Error GoTo 0

      If Not vbComp Is Nothing Then vbProj.VBComponents.Remove vbComp
      vbProj.VBComponents.Import modulePath
  End Sub

  Private Function ReadAllText(ByVal path As String) As String
      Dim fso As Object, ts As Object
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set ts = fso.OpenTextFile(path, 1)
      ReadAllText = ts.ReadAll
      ts.Close
  End Function

  Private Function ParseCsvLine(ByVal line As String) As Variant
      Dim items As Collection
      Set items = New Collection

      Dim i As Long, ch As String, inQuotes As Boolean, cur As String
      For i = 1 To Len(line)
          ch = Mid$(line, i, 1)
          If ch = """" Then
              If inQuotes And i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                  cur = cur & """"
                  i = i + 1
              Else
                  inQuotes = Not inQuotes
              End If
          ElseIf ch = "," And Not inQuotes Then
              items.Add cur
              cur = ""
          Else
              cur = cur & ch
          End If
      Next i
      items.Add cur

      Dim arr() As String
      ReDim arr(0 To items.count - 1)
      Dim idx As Long
      For idx = 1 To items.count
          arr(idx - 1) = items(idx)
      Next idx

      ParseCsvLine = arr
  End Function
