' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "ModuleColor"
Sub ColorByContext()
    'Macro principale : ColorByContext
    Dim wsActive As Worksheet
    Set wsActive = ThisWorkbook.ActiveSheet
    If wsActive.Name = "Roulement" Then
        ' Si l'onglet actif est Roulement
        Call ColorRoulement
    Else
        ' Si l'onglet actif est un mois spécifique
        Dim action As Integer
        action = MsgBox("Voulez-vous mettre à jour uniquement cet onglet ? (Non = Appliquer sur l'année)", vbYesNo + vbQuestion, "Choix")
        If action = vbYes Then
            Call ColorMonth
        Else
            Call ColorYear
        End If
    End If
End Sub

Sub ColorRoulement()
    ' colorie le planning de l'onglet roulement
    On Error GoTo ErrorHandler
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        .EnableEvents = False
    End With

    Dim wsActive As Worksheet
    Dim dict As Object
    Dim valuesSource As Variant
    Dim valuesTarget As Variant
    Dim colorData As Variant
    Dim sourceRange As Range
    Dim i As Long, j As Long
    Dim cellValue As Variant
    Set wsActive = ThisWorkbook.Sheets("Roulements")
    Set dict = CreateObject("Scripting.Dictionary")
    Set sourceRange = ThisWorkbook.Sheets("Config_Calendrier").Range("CP2:CP213")

    ' Charger les valeurs sources
    valuesSource = sourceRange.value

    ' Peupler le dictionnaire avec les couleurs et les styles
    For i = 1 To UBound(valuesSource, 1)
        cellValue = valuesSource(i, 1)
        If Not dict.Exists(cellValue) Then
            Set cellItem = sourceRange.Cells(i, 1)
            dict.Add cellValue, Array(cellItem.Interior.Color, cellItem.Font.Color)
        End If
    Next i

    ' Charger et traiter les données de destination
    With wsActive.Range("B6:BG31")
        valuesTarget = .value
        For i = 1 To UBound(valuesTarget, 1)
            For j = 1 To UBound(valuesTarget, 2)
                cellValue = valuesTarget(i, j)
                If dict.Exists(cellValue) Then
                    colorData = dict(cellValue)
                    .Cells(i, j).Interior.Color = colorData(0)
                    .Cells(i, j).Font.Color = colorData(1)
                End If

                ' Vérifier le code horaire complet
                If cellValue = "8:30 12:45 16:30 20:15" Then
                    With .Cells(i, j).Font
                        .Name = "Arial Narrow"
                        .Size = 8
                        .Bold = False
                        .Color = RGB(0, 0, 0) ' Noir
                    End With
                End If
            Next j
        Next i
    End With

CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur est survenue dans ColorRoulement : " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Sub ColorYear()
    'Macro pour les 12 mois : ColorYear
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wsSource As Worksheet
    Dim ws As Worksheet
    Dim dict As Object
    Dim cell As Range
    Dim lundiDemande As Date
    Dim planningRange As Range
    Dim firstMondayCol As Range
    Dim colStart As Long

    ' Demander la date du lundi de départ
    lundiDemande = InputBox("Entrez la date du lundi (jj/mm/aaaa) pour commencer la coloration :")
    If IsDate(lundiDemande) = False Then
        MsgBox "La date entrée n'est pas valide.", vbExclamation
        Exit Sub
    End If

    ' Initialiser les couleurs et dictionnaire
    Set dict = CreateObject("Scripting.Dictionary")
    Set wsSource = ThisWorkbook.Sheets("Config_Calendrier")
    For Each cell In wsSource.Range("CP2:CP213")
        If Not dict.Exists(cell.value) Then
            dict.Add cell.value, Array(cell.Interior.Color, cell.Font.Color)
        End If
    Next cell

    ' Colorer les onglets des mois
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Roulement" Then
            On Error Resume Next
            Set planningRange = Nothing
            Set planningRange = ws.Range("planning")
            On Error GoTo ErrorHandler
            If Not planningRange Is Nothing Then
                ' Identifier la colonne correspondant au lundi demandé
                Set firstMondayCol = ws.Rows(5).Find(What:=lundiDemande, LookIn:=xlValues, LookAt:=xlWhole)
                If Not firstMondayCol Is Nothing Then
                    colStart = firstMondayCol.Column
                    Dim rowStart As Long, rowEnd As Long, colEnd As Long
                    rowStart = 6
                    rowEnd = planningRange.Rows.Count + 5
                    colEnd = planningRange.Columns.Count

                    ' Colorer les cellules à partir de la colonne du lundi demandé
                    For Each cell In ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
                        If dict.Exists(cell.value) Then
                            cell.Interior.Color = dict(cell.value)(0)
                            cell.Font.Color = dict(cell.value)(1)
                        End If
                        
                        ' Vérifier le code horaire complet
                        If cell.value = "8:30 12:45 16:30 20:15" Then
                            With cell.Font
                                .Name = "Arial Narrow"
                                .Size = 8
                                .Bold = False
                                .Color = RGB(0, 0, 0) ' Noir
                            End With
                        End If
                    Next cell
                End If
            End If
        End If
    Next ws

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur est survenue dans ColorYear : " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Sub ColorMonth()
    'Macro pour un mois spécifique : ColorMonth
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wsActive As Worksheet
    Dim dict As Object
    Dim planningRange As Range
    Dim cell As Range

    ' Onglet actif
    Set wsActive = ThisWorkbook.ActiveSheet

    ' Initialiser les couleurs et dictionnaire
    Set dict = CreateObject("Scripting.Dictionary")
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Sheets("Config_Calendrier")
    For Each cell In wsSource.Range("CP2:CP213")
        If Not dict.Exists(cell.value) Then
            dict.Add cell.value, Array(cell.Interior.Color, cell.Font.Color)
        End If
    Next cell

    ' Définir la plage de planification
    Set planningRange = wsActive.Range("planning")

    ' Colorer les cellules
    For Each cell In planningRange
        If dict.Exists(cell.value) Then
            cell.Interior.Color = dict(cell.value)(0)
            cell.Font.Color = dict(cell.value)(1)
        End If
        
        ' Vérifier le code horaire complet
        If cell.value = "8:30 12:45 16:30 20:15" Then
            With cell.Font
                .Name = "Arial Narrow"
                .Size = 8
                .Bold = False
                .Color = RGB(0, 0, 0) ' Noir
            End With
        End If
    Next cell

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur est survenue dans ColorMonth : " & Err.Description, vbCritical
    Resume CleanUp
End Sub
