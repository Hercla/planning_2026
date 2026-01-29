VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo SafeExit

    ' (optionnel) audit config keys
    On Error Resume Next
    ModuleCFG_Audit.EnsureConfigKeys True
    On Error GoTo SafeExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Applique la vue uniquement selon le scope (ACTIVE/ALL) depuis config
    Module_ViewApply.VIEW_Apply_ByScope

SafeExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Module_CFG_Events.CFG_OnWorkbookClose_CancelPending
    On Error GoTo 0
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("AG41")) Is Nothing Then ' Vérifie si la cellule modifiée est AG41
        If Target.value = "x" Then
            Call UpdateF3     'appelle la macro UpdateF3 pour mettre à jour la valeur de F3
            Range("Us 1D 1").Copy 'copie les données à coller
            Sheets.Add After:=Sheets(Sheets.count) 'ajoute un nouvel onglet
            Range("Us 1D 1").PasteSpecial 'colle les données dans la cellule Us 1D 1 du nouvel onglet
        End If
    End If
End Sub

Sub UpdateF3()
    Dim mois As String
    mois = UCase(Format(Date, "mmm")) 'récupère le mois en cours
    Select Case mois 'sélectionne la valeur à mettre dans F3 en fonction du mois en cours
        Case "JANV"
            Sheets("Model").Range("F3").value = 1
        Case "FEV"
            Sheets("Model").Range("F3").value = 2
        Case "MARS"
            Sheets("Model").Range("F3").value = 3
        Case "AVRIL"
            Sheets("Model").Range("F3").value = 4
        Case "MAI"
            Sheets("Model").Range("F3").value = 5
        Case "JUIN"
            Sheets("Model").Range("F3").value = 6
        Case "JUIL"
            Sheets("Model").Range("F3").value = 7
        Case "AOUT"
            Sheets("Model").Range("F3").value = 8
        Case "SEPT"
            Sheets("Model").Range("F3").value = 9
        Case "OCT"
            Sheets("Model").Range("F3").value = 10
        Case "NOV"
            Sheets("Model").Range("F3").value = 11
        Case "DEC"
            Sheets("Model").Range("F3").value = 12
    End Select
End Sub
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    Dim aujourdhui As Date
    Dim jourActuel As Integer
    Dim moisActuel As Integer
    Dim anneeActuelle As Integer
    Dim anneePlanning As Long
    Dim nomMoisActuelOnglet As String
    Dim moisFrancais As Variant
    
    Dim plageDeRechercheJoursLigne4 As Range
    Dim plageIndicateurLigne2 As Range
    
    Dim cellDeRecherche As Range
    Dim cellPourIndicateurLigne2 As Range
    Dim i As Integer
    ' Dim originalValue As Variant ' Moins nécessaire avec la nouvelle logique de nettoyage pour ligne 2

    ' Indicateur est maintenant juste "Today"
    Const INDICATEUR_JOUR As String = "Today" ' <<<<< CHANGÉ POUR "Today" (sans espace après)

    Application.ScreenUpdating = False

    moisFrancais = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                         "Juillet", "Aout", "Sept", "Oct", "Nov", "Dec")

    aujourdhui = Date
    jourActuel = Day(aujourdhui)
    moisActuel = Month(aujourdhui)
    anneeActuelle = Year(aujourdhui)

    On Error Resume Next
    Dim anneePlanningVariant As Variant
    anneePlanningVariant = ThisWorkbook.Sheets("Config_Calendrier").Range("F22").value
    If Err.Number <> 0 Then
        Application.ScreenUpdating = True
        MsgBox "ERREUR: Impossible de lire l'année depuis 'Accueil'!F22. Erreur: " & Err.description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    If Not IsNumeric(anneePlanningVariant) Then
        Application.ScreenUpdating = True
        MsgBox "ATTENTION: L'année '" & anneePlanningVariant & "' en Acceuil!F22 n'est pas un nombre valide.", vbExclamation
        Exit Sub
    ElseIf CLng(anneePlanningVariant) < 1900 Or CLng(anneePlanningVariant) > 2100 Then
        Application.ScreenUpdating = True
        MsgBox "ATTENTION: L'année '" & anneePlanningVariant & "' en Acceuil!F22 n'est pas dans une plage valide (1900-2100).", vbExclamation
        Exit Sub
    End If
    anneePlanning = CLng(anneePlanningVariant)

    ' --- Nettoyage ---
    ' 1. Nettoyage de la Ligne 4 (plage de recherche des jours B4:AF4)
    '    Cette section est une précaution si un ancien script ou une manip manuelle a mis un indicateur ici.
    '    Normalement, avec ce script, la ligne 4 ne devrait contenir que les numéros de jours.
    On Error Resume Next
    Set plageDeRechercheJoursLigne4 = sh.Range("B4:AF4")
    If Err.Number = 0 Then
        For Each cellDeRecherche In plageDeRechercheJoursLigne4
            ' Si on trouve un ancien format "INDICATEUR_JOUR + numéro" sur la ligne 4
            If Left(CStr(cellDeRecherche.value), Len(INDICATEUR_JOUR) + 1) = INDICATEUR_JOUR & " " Then ' Ex: "Today 13"
                Dim valeurApresIndicateur As String
                valeurApresIndicateur = Trim(Mid(CStr(cellDeRecherche.value), Len(INDICATEUR_JOUR) + 2))
                If IsNumeric(valeurApresIndicateur) Then
                    cellDeRecherche.value = CInt(valeurApresIndicateur)
                Else
                    cellDeRecherche.value = valeurApresIndicateur ' Ou vider si ce n'est pas un nombre
                End If
                With cellDeRecherche.Font
                    .Bold = False
                    .ColorIndex = xlAutomatic
                End With
            ' Si on trouve juste "Today" sur la ligne 4 (improbable avec le fonctionnement normal)
            ElseIf CStr(cellDeRecherche.value) = INDICATEUR_JOUR Then
                 cellDeRecherche.value = "" ' Ou une valeur par défaut, car on ne sait pas quel jour restaurer
                 With cellDeRecherche.Font
                    .Bold = False
                    .ColorIndex = xlAutomatic
                End With
            End If
        Next cellDeRecherche
    Else
        Err.Clear
    End If
    On Error GoTo 0

    ' 2. Nettoyer la plage où l'INDICATEUR "Today" EST PLACÉ (Ligne 2 : B2:AF2)
    On Error Resume Next
    Set plageIndicateurLigne2 = sh.Range("B2:AF2")
    If Err.Number = 0 Then
        For Each cellPourIndicateurLigne2 In plageIndicateurLigne2
            If CStr(cellPourIndicateurLigne2.value) = INDICATEUR_JOUR Then ' Si la cellule contient "Today"
                Dim jourARestaurer As Variant
                ' Lire le numéro du jour depuis la ligne 4 de la même colonne
                jourARestaurer = sh.Cells(4, cellPourIndicateurLigne2.Column).value
                
                If IsNumeric(jourARestaurer) Then
                    cellPourIndicateurLigne2.value = CInt(jourARestaurer)
                Else
                    ' Si la ligne 4 ne contient pas un numéro pour cette colonne,
                    ' on pourrait vider la cellule ou mettre une valeur par défaut.
                    ' Pour l'instant, on la vide si ce n'est pas un nombre.
                    cellPourIndicateurLigne2.value = ""
                End If
                
                With cellPourIndicateurLigne2.Font ' Réinitialiser la police
                    .Bold = False
                    .ColorIndex = xlAutomatic
                End With
            End If
        Next cellPourIndicateurLigne2
    Else
        Err.Clear
    End If
    On Error GoTo 0
    
    ' --- Vérifications et Logique Principale ---
    If anneePlanning <> anneeActuelle Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim isMonthlySheet As Boolean
    isMonthlySheet = False
    For i = LBound(moisFrancais) To UBound(moisFrancais)
        If sh.Name = moisFrancais(i) Then
            isMonthlySheet = True
            Exit For
        End If
    Next i

    If Not isMonthlySheet Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    nomMoisActuelOnglet = moisFrancais(moisActuel - 1)

    If sh.Name = nomMoisActuelOnglet Then
        Dim jourTrouveSurLigne4 As Boolean
        jourTrouveSurLigne4 = False
        Dim colonneDuJour As Long
        colonneDuJour = 0

        Set plageDeRechercheJoursLigne4 = sh.Range("B4:AF4")
        For Each cellDeRecherche In plageDeRechercheJoursLigne4
            If Not IsEmpty(cellDeRecherche.value) Then
                If IsNumeric(Trim(CStr(cellDeRecherche.value))) Then
                    If CInt(Trim(CStr(cellDeRecherche.value))) = jourActuel Then
                        colonneDuJour = cellDeRecherche.Column
                        jourTrouveSurLigne4 = True
                        Exit For
                    End If
                End If
            End If
        Next cellDeRecherche

        If jourTrouveSurLigne4 And colonneDuJour > 0 Then
            Dim cellCibleLigne2 As Range
            Set cellCibleLigne2 = sh.Cells(2, colonneDuJour)
            
            ' Écrire seulement "Today"
            cellCibleLigne2.value = INDICATEUR_JOUR ' <<<<< CHANGÉ POUR ÉCRIRE SEULEMENT L'INDICATEUR
            
            With cellCibleLigne2.Font
                .Bold = True
                .Color = vbRed
            End With
        End If
    End If

    Application.ScreenUpdating = True
End Sub

