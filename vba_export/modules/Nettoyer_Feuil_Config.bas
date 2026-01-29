Attribute VB_Name = "Nettoyer_Feuil_Config"
 Option Explicit

  Public Sub Nettoyer_Feuil_Config()
      Dim ws As Worksheet
      Set ws = ThisWorkbook.Worksheets("Feuil_Config")

      Dim lastRow As Long
      lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

      Dim seen As Object
      Set seen = CreateObject("Scripting.Dictionary")
      seen.CompareMode = vbTextCompare

      Dim r As Long
      Application.ScreenUpdating = False

      For r = lastRow To 1 Step -1
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
      MsgBox "Feuil_Config nettoyée.", vbInformation
  End Sub
