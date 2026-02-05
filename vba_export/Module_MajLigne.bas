Attribute VB_Name = "Module_MajLigne"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

Public Sub MajLigne()
    Dim staffKey As String
    Dim ws As Worksheet

    Set ws = ActiveSheet

    ' Ouvrir le UserForm3 et récupérer le choix via Tag
    UserForm3.tag = vbNullString
    UserForm3.Show vbModal

    staffKey = Trim$(UserForm3.tag)
    Unload UserForm3

    If staffKey = vbNullString Then Exit Sub

    ' Afficher uniquement la ligne de la personne sélectionnée sur la feuille active
    ShowOnlyStaffLine ws, staffKey
End Sub

Private Sub ShowOnlyStaffLine(ByVal ws As Worksheet, ByVal staffKey As String)
    Dim firstDataRow As Long, lastDataRow As Long
    Dim r As Range

    If ws Is Nothing Then Exit Sub
    If Len(Trim$(staffKey)) = 0 Then Exit Sub

    ' Détecter le bloc "personnel" : on cherche l'en-tête "Nom" en colonne A
    firstDataRow = FindRowBelowHeader(ws, "Nom", 1)
    If firstDataRow = 0 Then firstDataRow = 2

    lastDataRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastDataRow < firstDataRow Then lastDataRow = firstDataRow

    ' Chercher la ligne (col A) : "Nom Prénom"
    Set r = ws.Columns(1).Find(What:=staffKey, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    ' Fallback si col A est "Nom_Prénom"
    If r Is Nothing Then
        Set r = ws.Columns(1).Find(What:=Replace(staffKey, " ", "_"), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    End If

    If r Is Nothing Then
        MsgBox "Introuvable sur la feuille """ & ws.Name & """ : " & staffKey, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Masquer toutes les lignes du personnel et réafficher la bonne ligne
    ws.Rows(firstDataRow & ":" & lastDataRow).Hidden = True
    ws.Rows(r.row).Hidden = False

    ' Amener la ligne à l'écran
    Application.GoTo ws.Cells(r.row, 1), True

    Application.ScreenUpdating = True
End Sub

Private Function FindRowBelowHeader(ByVal ws As Worksheet, ByVal headerText As String, ByVal headerCol As Long) As Long
    Dim f As Range
    Set f = ws.Columns(headerCol).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        FindRowBelowHeader = 0
    Else
        FindRowBelowHeader = f.row + 1
    End If
End Function

Public Sub ResetAfficherTout()
    Dim ws As Worksheet
    Dim firstDataRow As Long, lastDataRow As Long

    Set ws = ActiveSheet
    firstDataRow = FindRowBelowHeader(ws, "Nom", 1)
    If firstDataRow = 0 Then firstDataRow = 2
    lastDataRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    Application.ScreenUpdating = False
    ws.Rows(firstDataRow & ":" & lastDataRow).Hidden = False
    Application.ScreenUpdating = True
End Sub

