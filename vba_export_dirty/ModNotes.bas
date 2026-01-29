Attribute VB_Name = "ModNotes"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
' ===============================================
' === CODE À METTRE DANS UN MODULE STANDARD     ===
' ===============================================
Sub CreerNoteRemplacement()
    Dim targetCell As Range
    Dim form As UserForm4
    Dim personName As String
    Dim commentContent As String
    Dim fullCommentText As String

    If TypeName(Selection) <> "Range" Or Selection.Cells.CountLarge > 1 Then
        MsgBox "Veuillez sélectionner une seule cellule.", vbExclamation, "Sélection Invalide"
        Exit Sub
    End If
    Set targetCell = Selection

    Set form = New UserForm4
    form.Show
    
    If form.WasCancelled Then
        Unload form
        Set form = Nothing
        Exit Sub
    End If

    ' --- Récupérer les données du formulaire (AVEC LE NOUVE NOM) ---
    personName = form.cmbNom.value
    commentContent = form.txtCommentaire.value ' <-- LA MODIFICATION EST ICI
    
    Unload form
    Set form = Nothing

    ' --- Gestion du commentaire ---
    On Error GoTo ErrorHandler

    If Not targetCell.comment Is Nothing Then
        targetCell.comment.Delete
    End If

    If Trim(commentContent) <> "" Then
        fullCommentText = personName & ":" & vbCrLf & commentContent
    Else
        fullCommentText = personName
    End If
    
    targetCell.AddComment text:=fullCommentText
    
    With targetCell.comment.Shape
        .TextFrame.AutoSize = True
        .Visible = True
    End With

    targetCell.Select
    Exit Sub

ErrorHandler:
    MsgBox "Impossible de modifier la note." & vbCrLf & _
           "Vérifiez si la feuille est protégée.", vbCritical, "Erreur"
End Sub

