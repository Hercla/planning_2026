VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mai11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet0_Change(ByVal target As Range)
    ' Spécifie le contenu spécifique et la largeur souhaitée
    Dim specificCode As String
    Dim specificWidth As Double

    specificCode = "8:30 12:45 16:30 20:15"
    specificWidth = 4.9

    ' Vérifie si le changement concerne une seule cellule
    If target.Cells.count = 1 Then
        ' Vérifie si la cellule contient le code spécifique
        If target.value = specificCode Then
            ' Ajuste la largeur de la colonne à 4.90
            target.EntireColumn.ColumnWidth = specificWidth
        End If
    End If
End Sub
Private Sub CommandButton120_Click()
    UserForm3.Show
End Sub


