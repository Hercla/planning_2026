Attribute VB_Name = "ImportSelectedVBA"
  Option Explicit

  Public Sub ImportVBA_FromDialog()
      Dim fd As FileDialog
      Dim proj As Object
      Dim fso As Object
      Dim targetFolder As String

      targetFolder = "C:\Users\hercl\planning_2026\"

      On Error Resume Next
      ChDrive Left$(targetFolder, 2)
      ChDir targetFolder
      On Error GoTo 0

      Set fso = CreateObject("Scripting.FileSystemObject")
      Set proj = ThisWorkbook.VBProject

      Set fd = Application.FileDialog(msoFileDialogFilePicker)
      With fd
          .Title = "Choisir les fichiers VBA à importer"
          .AllowMultiSelect = True
          .Filters.Clear
          .Filters.Add "VBA Files", "*.bas;*.cls;*.frm", 1
          .InitialFileName = targetFolder
          If .Show <> -1 Then Exit Sub
      End With

      Application.ScreenUpdating = False
      Application.DisplayAlerts = False

      Dim i As Long, filePath As String, compName As String
      For i = 1 To fd.SelectedItems.count
          filePath = fd.SelectedItems(i)
          compName = fso.GetBaseName(filePath)

          RemoveComponentByName proj, compName
          proj.VBComponents.Import filePath
      Next i

      Application.DisplayAlerts = True
      Application.ScreenUpdating = True

      MsgBox "Import terminé.", vbInformation
  End Sub

  Private Sub RemoveComponentByName(ByVal proj As Object, ByVal compName As String)
      Dim comp As Object
      For Each comp In proj.VBComponents
          If StrComp(comp.Name, compName, vbTextCompare) = 0 Then
              If comp.Type <> 100 Then ' 100 = vbext_ct_Document (ThisWorkbook/Sheets)
                  proj.VBComponents.Remove comp
              End If
              Exit For
          End If
      Next comp
  End Sub
