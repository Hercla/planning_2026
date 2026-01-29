Attribute VB_Name = "Module_Macros_Génère_copy_model"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

Sub Remplir_dates(Optional yearInput As Long = 0)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Vous pouvez spécifier une feuille spécifique si nécessaire

    Dim j As Long, k As Long
    Dim mois As Integer
    Dim date_debut As Date, date_fin As Date

    ' Supprime la couleur dans les lignes précédemment colorées
    For j = 7 To 37
        With ws.Range("A" & j)
            If .Interior.ColorIndex = 37 Then
                ws.Range("A" & j & ":B" & j & ",D" & j & ":F" & j).Interior.ColorIndex = xlNone
            End If
        End With
    Next j

    ' Récupère le numéro du mois depuis la cellule F3
    mois = ws.Range("F3").value
    If mois < 1 Or mois > 12 Then
        MsgBox "Le numéro du mois dans F3 est invalide : " & mois, vbExclamation
        GoTo CleanUp
    End If

    ' Définit l'année
    If yearInput = 0 Then yearInput = Year(Date)

    ' Définit la date de début et de fin du mois
    date_debut = DateSerial(yearInput, mois, 1)
    date_fin = DateSerial(yearInput, mois + 1, 0)

    k = 7 ' Ligne de départ pour remplir les dates

    ' Remplit les dates dans les colonnes A et B
    While date_debut <= date_fin And k <= 37
        ws.Range("A" & k).value = Weekday(date_debut, vbSunday) ' 1 = Dimanche, 7 = Samedi
        ws.Range("B" & k).value = date_debut
        date_debut = date_debut + 1
        k = k + 1
    Wend

    ' Change la couleur des lignes correspondant à des samedis ou dimanches
    For j = 7 To 37
        With ws.Range("A" & j)
            If .value = vbSunday Or .value = vbSaturday Then
                ws.Range("A" & j & ":B" & j & ",D" & j & ":F" & j).Interior.Color = RGB(204, 229, 255)
            End If
        End With
    Next j

CleanUp:
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur est survenue dans Remplir_dates : " & Err.description, vbCritical
    Resume CleanUp
End Sub
Sub UpdateF3()
'en fonction de l'onglet actif le converti en chiffre et le colle dans le fichier destination
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    Dim num As Integer
    
    Select Case ws.Name
        Case "Janv"
            num = 1
        Case "Fev"
            num = 2
        Case "Mars"
            num = 3
        Case "Avril"
            num = 4
        Case "Mai"
            num = 5
        Case "Juin"
            num = 6
        Case "Juillet"
            num = 7
        Case "Aout"
            num = 8
        Case "Sept"
            num = 9
        Case "Oct"
            num = 10
        Case "Nov"
            num = 11
        Case "Dec"
            num = 12
        Case Else
            Exit Sub ' Quitte la macro si le nom de l'onglet n'est pas reconnu
    End Select
    
    Dim destinationWorkbook As Workbook
    Set destinationWorkbook = Workbooks.Open(fileName:="C:\Users\claud\OneDrive\IC.1 Admin & Team\1. Hot Topics\Horaire_2023\Générer Demandes Remplacements.xlsm")
    
    ' Modifie la valeur de la cellule F3 dans l'onglet "Model" du classeur de destination
    destinationWorkbook.Worksheets("Model").Range("F3").value = num
    
    ' Appelle la macro "OpenNewTabsWithModelData" dans le fichier de destination
    destinationWorkbook.Application.Run "'" & destinationWorkbook.Name & "'!OpenNewTabsWithModelData"
    
    ' Réactive le fichier source
    ws.Activate
                        
End Sub
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' Vérifie si la modification a eu lieu dans la cellule F3
    If Not Intersect(Target, Me.Range("F3")) Is Nothing Then
        ' Désactive les événements pour éviter les appels récursifs
        Application.EnableEvents = False
        Call Remplir_dates
    End If

CleanUp:
    ' Réactive les événements
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur est survenue dans l'événement Worksheet_Change : " & Err.description, vbCritical
    Resume CleanUp
End Sub

Sub ColorCells2()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim codesRange As Range
    Dim planningRange As Range
    Dim dict As Object
    Dim cell As Range
    Dim darkGreenColor As Long
    
    ' Définir le classeur source
    Dim sourceWorkbookName As String
    sourceWorkbookName = "Planning_2025.xlsm"
    
    ' Vérifie si le classeur source est ouvert
    Dim wbSource As Workbook
    On Error Resume Next
    Set wbSource = Workbooks(sourceWorkbookName)
    On Error GoTo ErrorHandler
    
    If wbSource Is Nothing Then
        ' Ouvre le classeur source si ce n'est pas le cas
        Dim sourcePath As String
        sourcePath = "C:\Chemin\Vers\Votre\Classeur\" & sourceWorkbookName ' Modifiez le chemin selon vos besoins
        If Dir(sourcePath) = "" Then
            MsgBox "Le classeur source n'a pas été trouvé : " & sourcePath, vbCritical
            GoTo CleanUp
        End If
        Set wbSource = Workbooks.Open(fileName:=sourcePath, ReadOnly:=True)
    End If
    
    ' Définir la feuille source contenant les codes
    On Error Resume Next
    Set wsSource = wbSource.Sheets("Config_Calendrier")
    On Error GoTo ErrorHandler
    If wsSource Is Nothing Then
        MsgBox "La feuille 'Acceuil' n'existe pas dans " & sourceWorkbookName, vbCritical
        GoTo CleanUp
    End If
    
    ' Définir la feuille de destination (active sheet ou spécifiée)
    Set wsDest = ThisWorkbook.ActiveSheet ' Vous pouvez spécifier une feuille spécifique si nécessaire
    
    ' Définir les plages
    Set codesRange = wsSource.Range("CP2:CP213")
    On Error Resume Next
    Set planningRange = wsDest.Range("planning")
    On Error GoTo ErrorHandler
    If planningRange Is Nothing Then
        MsgBox "La plage nommée 'planning' n'est pas définie dans la feuille " & wsDest.Name, vbCritical
        GoTo CleanUp
    End If
    
    ' Initialiser le dictionnaire
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Couleur vert foncé
    darkGreenColor = RGB(0, 100, 0)
    
    ' Stocke les valeurs des couleurs dans un dictionnaire pour des recherches plus rapides
    Dim codeValue As Variant
    For Each cell In codesRange
        If Not dict.Exists(cell.value) Then
            dict.Add cell.value, Array(cell.Interior.Color, cell.Font.Color)
        End If
    Next cell
    
    ' Traite les cellules de la plage de planification
    For Each cell In planningRange
        ' Ne modifie pas les cellules déjà colorées en vert foncé
        If cell.Interior.Color <> darkGreenColor Then
            If dict.Exists(cell.value) Then
                cell.Interior.Color = dict(cell.value)(0)
                cell.Font.Color = dict(cell.value)(1)
                cell.Font.Name = "Arial"
                cell.Font.Size = 12
                cell.Font.Bold = False
            End If
        End If
    Next cell
    
CleanUp:
    ' Ferme le classeur source si il a été ouvert par la macro
    If Not wbSource Is Nothing Then
        If wbSource.ReadOnly Then
            wbSource.Close SaveChanges:=False
        End If
    End If
    
    ' Réactive les paramètres d'Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Une erreur est survenue dans ColorCells2 : " & Err.description, vbCritical
    Resume CleanUp
End Sub

Sub ClearAndWhitenCell()
    Dim selectedCells As Range
    
    ' Vérifie si des cellules sont sélectionnées
    If Not Application.Selection Is Nothing Then
        ' Vérifie que la sélection est bien des cellules
        If TypeName(Application.Selection) = "Range" Then
            Set selectedCells = Application.Selection
            
            ' Efface le contenu des cellules sélectionnées
            selectedCells.ClearContents
            
            ' Remet le fond des cellules sélectionnées en blanc
            selectedCells.Interior.Color = RGB(255, 255, 255)
        Else
            MsgBox "Veuillez sélectionner des cellules avant d'exécuter cette macro.", vbExclamation
        End If
    Else
        MsgBox "Aucune cellule sélectionnée.", vbExclamation
    End If
End Sub

Sub FormatCell(cell As Range)
    On Error GoTo ErrorHandler
    
    If cell Is Nothing Then
        MsgBox "La cellule spécifiée est invalide.", vbExclamation
        Exit Sub
    End If
    
    If cell.value = "WE" Then
        cell.Font.Color = RGB(255, 255, 255) ' Blanc
        cell.Interior.Color = RGB(0, 0, 255) ' Bleu
    Else
        ' Remet les couleurs par défaut si la valeur n'est pas "WE"
        cell.Interior.Color = RGB(255, 255, 255) ' Blanc
        cell.Font.Color = RGB(0, 0, 0) ' Noir
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Une erreur est survenue dans FormatCell : " & Err.description, vbCritical
End Sub
' Déclaration d'une variable publique pour stocker la plage copiée
Public CopiedRange As Range










