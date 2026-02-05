' ExportedAt: 2026-01-12 15:37:08 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "ColorShortages"
Sub ColorShortages_ActiveSheet_BasedOnColor()

    Dim ws As Worksheet
    Dim col As Long
    Dim colorIndexDay As Long
    
    ' Variables de seuil
    Dim thresholdMorning As Long
    Dim thresholdAfternoon As Long
    Dim thresholdEvening As Long
    
    Set ws = ActiveSheet
    
    On Error GoTo ErrHandler
    
    For col = 2 To 32
        
        ' Récupérer l'Index de couleur de la cellule en ligne 3
        colorIndexDay = ws.Cells(3, col).Interior.ColorIndex
        
        ' S'il est rouge (généralement 3) => Sam/Dim/Férié
        If colorIndexDay = 3 Then
            thresholdMorning = 5
            thresholdAfternoon = 2
            thresholdEvening = 3
        Else
            thresholdMorning = 7
            thresholdAfternoon = 3
            thresholdEvening = 3
        End If
        
        ' Idem que précédemment
        If ws.Cells(60, col).value < thresholdMorning Then
            ws.Cells(60, col).Interior.Color = RGB(255, 199, 206)
        Else
            ws.Cells(60, col).Interior.ColorIndex = xlNone
        End If
        
        If ws.Cells(61, col).value < thresholdAfternoon Then
            ws.Cells(61, col).Interior.Color = RGB(255, 199, 206)
        Else
            ws.Cells(61, col).Interior.ColorIndex = xlNone
        End If
        
        If ws.Cells(62, col).value < thresholdEvening Then
            ws.Cells(62, col).Interior.Color = RGB(255, 199, 206)
        Else
            ws.Cells(62, col).Interior.ColorIndex = xlNone
        End If
    
    Next col
    
    MsgBox "Coloration terminée pour la feuille : " & ws.Name & " !"
    Exit Sub

ErrHandler:
    MsgBox "Une erreur s'est produite : " & Err.Description

End Sub

