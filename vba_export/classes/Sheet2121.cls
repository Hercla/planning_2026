VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton120_Click()
    UserForm3.Show
End Sub
' === MODULE FEUILLE "Roulements" ===

Private Sub Worksheet_Change(ByVal target As Range)
    On Error GoTo fin
    ' Ne déclenche la logique que si la modification est dans la plage des horaires
    If Not Intersect(target, Range("D6:BG31")) Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        Call RecopieCyclique(target)

fin:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

' Recopie le code horaire cycliquement selon la valeur colonne C (cycle 1 à 4)
Private Sub RecopieCyclique(target As Range)
    Dim cyc As Long, codeHoraire As Variant
    Dim colDebut As Long, colFin As Long, colActuelle As Long

    cyc = Cells(target.row, "C").value
    codeHoraire = target.value
    If cyc >= 1 And cyc <= 4 Then
        FormatCelluleHoraire target, codeHoraire, target
        colDebut = target.Column
        colFin = Range("BG1").Column
        colActuelle = colDebut + (cyc * 7)
        Do While colActuelle <= colFin
            Cells(target.row, colActuelle).value = codeHoraire
            FormatCelluleHoraire Cells(target.row, colActuelle), codeHoraire, target
            colActuelle = colActuelle + (cyc * 7)
        Loop
    End If
End Sub

' Formate la cellule selon le code horaire (police, taille, couleur)
Private Sub FormatCelluleHoraire(ByRef cell As Range, ByVal codeHoraire As String, ByVal sourceCell As Range)
    If codeHoraire = "8:30 12:45 16:30 20:15" Then
        cell.Font.Name = "Arial Narrow"
        cell.Font.Size = 8
    Else
        cell.Font.Name = "Arial"
        cell.Font.Size = 12
        cell.Font.Bold = False
    End If
    If sourceCell.Interior.Color = vbYellow Then cell.Interior.Color = vbYellow
End Sub


