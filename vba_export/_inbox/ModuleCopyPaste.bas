' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "ModuleCopyPaste"
Option Explicit

' --- CopyForPasting Macro ---
Sub CopyForPasting()
    ' Vérifie si une sélection est faite
    If Not Application.Selection Is Nothing Then
        ' Assure que la sélection est une plage de cellules
        If TypeName(Application.Selection) = "Range" Then
            Set CopiedRange = Application.Selection
            MsgBox "Sélection copiée. Cliquez sur une cellule dans la plage 'planning' et exécutez PasteToPlanning pour coller.", vbInformation
        Else
            MsgBox "Veuillez sélectionner une plage de cellules à copier.", vbExclamation
        End If
    Else
        MsgBox "Aucune cellule sélectionnée à copier.", vbExclamation
    End If
End Sub

' --- PasteToPlanning Macro ---
Sub PasteToPlanning()
    ' Vérifie si CopiedRange est défini
    If Not CopiedRange Is Nothing Then
        Dim wsDest As Worksheet
        Set wsDest = ThisWorkbook.ActiveSheet ' Vous pouvez spécifier une feuille spécifique si nécessaire

        ' Vérifie si la cellule active est dans la plage nommée "planning"
        If Not Intersect(ActiveCell, wsDest.Range("planning")) Is Nothing Then
            ' Colle la plage copiée à l'emplacement de la cellule active
            CopiedRange.Copy
            ActiveCell.PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False ' Efface le presse-papiers
            MsgBox "Contenu collé avec succès.", vbInformation
        Else
            MsgBox "La cellule active n'est pas dans la plage 'planning'.", vbExclamation
        End If
    Else
        MsgBox "Aucune plage copiée. Utilisez CopyForPasting d'abord.", vbExclamation
    End If
End Sub

