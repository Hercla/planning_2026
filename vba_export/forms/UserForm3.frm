VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   2265
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4571
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Personnel")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    ' Remplir les ComboBox
    Me.cmbName.Clear
    Me.cmbFirstName.Clear
    For i = 2 To lastRow
        Me.cmbName.AddItem ws.Cells(i, 2).value ' Nom
        Me.cmbFirstName.AddItem ws.Cells(i, 3).value ' Prénom
    Next i
End Sub

Private Sub cmbName_Change()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Personnel")
    For i = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
        If ws.Cells(i, 2).value = Me.cmbName.value Then
            Me.cmbFirstName.value = ws.Cells(i, 3).value
            Exit For
        End If
    Next i
End Sub

Private Sub cmbFirstName_Change()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Personnel")
    For i = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
        If ws.Cells(i, 3).value = Me.cmbFirstName.value Then
            Me.cmbName.value = ws.Cells(i, 2).value
            Exit For
        End If
    Next i
End Sub

Private Sub btnSubmit_Click()
    ' Définir le nom sélectionné dans la propriété Tag du UserForm
    Me.tag = Me.cmbName.value & " " & Me.cmbFirstName.value
    Me.Hide
End Sub


