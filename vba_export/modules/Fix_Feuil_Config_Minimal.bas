Attribute VB_Name = "Fix_Feuil_Config_Minimal"
Option Explicit

  Public Sub Fix_Feuil_Config_Minimal()
      Dim ws As Worksheet
      Set ws = ThisWorkbook.Worksheets("Feuil_Config")

      Application.ScreenUpdating = False

      ' Supprime colonnes inutiles (C:I)
      ws.Range("C:I").Delete

      ' Titres propres
      ws.Cells(1, 1).value = "Column1"
      ws.Cells(1, 2).value = "Column2"

      ' Nettoyage doublons
      Dim lastRow As Long, r As Long
      Dim seen As Object
      Set seen = CreateObject("Scripting.Dictionary")
      seen.CompareMode = vbTextCompare

      lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
      For r = lastRow To 2 Step -1
          Dim key As String
          key = Trim$(CStr(ws.Cells(r, 1).value))
          If Len(key) = 0 Then
              ' ignore
          ElseIf seen.Exists(key) Then
              ws.Rows(r).Delete
          Else
              seen.Add key, True
          End If
      Next r

      Application.ScreenUpdating = True
      MsgBox "Feuil_Config nettoyée (A/B uniquement).", vbInformation
  End Sub
 Option Explicit

  Public Sub CheckMissingConfigKeys()
      Dim wsCfg As Worksheet
      Set wsCfg = ThisWorkbook.Worksheets("Feuil_Config")

      Dim keys As Object
      Set keys = CreateObject("Scripting.Dictionary")
      keys.CompareMode = vbTextCompare

      Dim lastRow As Long, r As Long, k As String
      lastRow = wsCfg.Cells(wsCfg.Rows.count, 1).End(xlUp).row
      For r = 2 To lastRow
          k = Trim$(CStr(wsCfg.Cells(r, 1).value))
          If Len(k) > 0 Then keys(k) = True
      Next r

      Dim missing As Object
      Set missing = CreateObject("Scripting.Dictionary")
      missing.CompareMode = vbTextCompare

      Dim comp As Object, code As String
      For Each comp In ThisWorkbook.VBProject.VBComponents
          code = ""
          On Error Resume Next
          code = comp.CodeModule.lines(1, comp.CodeModule.CountOfLines)
          On Error GoTo 0

          If Len(code) > 0 Then
              ExtractMissingKeys code, keys, missing
          End If
      Next comp

      Dim report As String
      report = "Clés manquantes dans Feuil_Config:" & vbCrLf & vbCrLf

      If missing.count = 0 Then
          report = report & "Aucune clé manquante ?"
      Else
          Dim key As Variant
          For Each key In missing.keys
              report = report & "- " & key & vbCrLf
          Next key
      End If

      MsgBox report, vbInformation
  End Sub

  Private Sub ExtractMissingKeys(ByVal code As String, ByVal keys As Object, ByVal missing As Object)
      Dim re As Object, matches As Object, m As Object
      Set re = CreateObject("VBScript.RegExp")
      re.Global = True
      re.IgnoreCase = True

      ' Capture: CfgText("KEY") / CfgValueOr("KEY", ...) etc.
      re.pattern = "Cfg(Text|Value|TextOr|ValueOr|Long|LongOr|Bool)\s*\(\s*""([^""]+)"""
      Set matches = re.Execute(code)

      For Each m In matches
          Dim k As String
          k = Trim$(m.SubMatches(1))
          If Len(k) > 0 Then
              If Not keys.Exists(k) Then
                  missing(k) = True
              End If
          End If
      Next m
  End Sub
