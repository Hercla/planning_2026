' ExportedAt: 2026-01-12 15:37:08 | Workbook: Planning_2026.xlsm
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageWorkTimeForm 
   Caption         =   "UserForm3"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4802
   OleObjectBlob   =   "ManageWorkTimeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManageWorkTimeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnStartDate_Click()
    ToggleDatePicker Me, Me.Controls("txtStartDate")
End Sub

Private Sub btnEndDate_Click()
    ToggleDatePicker Me, Me.Controls("txtEndDate")
End Sub

Private Sub cmbNom_Change()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Personnel")
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 2).value = Me.Controls("cmbNom").value Then
            Me.Controls("cmbPrenom").value = ws.Cells(i, 3).value
            Exit For
        End If
    Next i
End Sub

Private Sub cmbPrenom_Change()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Personnel")
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 3).value = Me.Controls("cmbPrenom").value Then
            Me.Controls("cmbNom").value = ws.Cells(i, 2).value
            Exit For
        End If
    Next i
End Sub

Private Sub UserForm_Initialize()
    FillEmployeeComboBoxes
End Sub

