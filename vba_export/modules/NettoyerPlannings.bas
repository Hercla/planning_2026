Attribute VB_Name = "NettoyerPlannings"
' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
Sub NettoyerPlanningsCouleursEtNotes()
    Dim mois As Variant, ws As Worksheet
    Dim dernièreLigne As Long, dernièreColonne As Long
    Dim zone As Range
    Dim shp As Shape, shpRange As Range

    'Feuilles mensuelles à nettoyer
    mois = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                 "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")

    For Each ws In ThisWorkbook.Worksheets
        If Not IsError(Application.Match(ws.Name, mois, 0)) Then
            'Détermine la zone de planning (C6 jusqu’à la dernière ligne/colonne utilisée)
            dernièreLigne = ws.Cells(ws.Rows.count, "A").End(xlUp).row
            dernièreColonne = ws.Cells(4, ws.Columns.count).End(xlToLeft).Column
            Set zone = ws.Range(ws.Cells(6, 3), ws.Cells(dernièreLigne, dernièreColonne))

            'Supprimer les valeurs saisies (pas les formules)
            On Error Resume Next
            zone.SpecialCells(xlCellTypeConstants).ClearContents
            On Error GoTo 0

            'Enlever les couleurs de fond dans la zone
            zone.Interior.ColorIndex = xlColorIndexNone

            'Supprimer les formes (rectangles/auto-shapes) superposées à la zone
            For Each shp In ws.Shapes
                If shp.Type <> msoChart And shp.Type <> msoComment Then
                    On Error Resume Next
                    Set shpRange = shp.TopLeftCell
                    On Error GoTo 0
                    If Not shpRange Is Nothing Then
                        If Not Intersect(shpRange, zone) Is Nothing Then shp.Delete
                    End If
                End If
            Next shp

            'Supprimer les commentaires et notes des cellules de la zone
            'Les méthodes ClearComments et ClearNotes effacent respectivement les anciens commentaires et les notes modernes:contentReference[oaicite:2]{index=2}.
            On Error Resume Next
            zone.ClearComments
            zone.ClearNotes
            On Error GoTo 0
        End If
    Next ws

    MsgBox "Les plannings, leurs couleurs, leurs rectangles et toutes les notes/commentaires ont été supprimés."
End Sub


