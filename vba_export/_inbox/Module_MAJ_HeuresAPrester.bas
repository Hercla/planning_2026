' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module_MAJ_HeuresAPrester"
Option Explicit

'===================================================================================
' MODULE :      Module_SyntheseMensuelle (Version Finale Robuste)
' DESCRIPTION : Met à jour les "Heures à prester" de manière optimisée en se
'               basant sur la feuille "Personnel". Gère les erreurs de nom d'onglet.
'===================================================================================

'--- Constantes pour rendre le code plus lisible et facile à maintenir ---
Private Const PERSONNEL_SHEET As String = "Personnel"
Private Const ACCUEIL_SHEET As String = "Accueil"
Private Const YEAR_CELL As String = "F22"
Private Const START_ROW As Long = 6

Public Sub MAJ_HeuresAPrester()
    '--- Déclaration des variables ---
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsPlan As Worksheet, wsPers As Worksheet, wsAccueil As Worksheet
    Dim mois As String, annee As Integer
    Dim lastRowPlan As Long, lastRowPers As Long
    Dim colNomPlan As Long, colMatPlan As Long, colHeuresAPrester As Long
    Dim colMoisPers As Long, colMatPers As Long
    Dim personnelDict As Object
    Dim arrPlan As Variant, arrPers As Variant, arrResultats As Variant
    Dim i As Long
    
    '--- Optimisations de performance ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '--- ÉTAPE 1: Préparation et validation robuste des feuilles ---
    ' Cette boucle est plus tolérante aux espaces dans les noms d'onglets
    Dim foundPers As Boolean, foundAccueil As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Trim(ws.Name) = PERSONNEL_SHEET Then
            Set wsPers = ws
            foundPers = True
        End If
        If Trim(ws.Name) = ACCUEIL_SHEET Then
            Set wsAccueil = ws
            foundAccueil = True
        End If
    Next ws

    If Not foundPers Or Not foundAccueil Then
        MsgBox "ERREUR : La feuille '" & PERSONNEL_SHEET & "' ou '" & ACCUEIL_SHEET & "' est introuvable. " & _
               "Veuillez vérifier les noms des onglets.", vbCritical
        GoTo CleanUp
    End If
    
    annee = wsAccueil.Range(YEAR_CELL).value
    If Not IsNumeric(annee) Or annee < 2000 Then
        MsgBox "Année du planning introuvable ou incorrecte dans " & ACCUEIL_SHEET & "!" & YEAR_CELL, vbCritical
        GoTo CleanUp
    End If
    
    '--- ÉTAPE 2: Trouver les colonnes nécessaires (de manière robuste) ---
    Set wsPlan = ActiveSheet
    colNomPlan = TrouverColonne(wsPlan, "Nom")
    colMatPlan = TrouverColonne(wsPlan, "Matricule")
    colHeuresAPrester = TrouverColonne(wsPlan, "Heures à prester")
    colMatPers = 1 ' On suppose que le matricule est toujours en colonne A de "Personnel"
    
    If colNomPlan = 0 Or colMatPlan = 0 Or colHeuresAPrester = 0 Then
        MsgBox "Une colonne clé ('Nom', 'Matricule' ou 'Heures à prester') n'a pas été trouvée sur le planning !", vbCritical
        GoTo CleanUp
    End If
    
    ' Trouver la colonne du mois dans "Personnel" (ex: "Janv %")
    mois = wsPlan.Name
    For i = 1 To wsPers.Cells(1, wsPers.Columns.Count).End(xlToLeft).Column
        If InStr(1, wsPers.Cells(1, i).value, mois, vbTextCompare) > 0 And InStr(1, wsPers.Cells(1, i).value, "%") > 0 Then
            colMoisPers = i
            Exit For
        End If
    Next i

    If colMoisPers = 0 Then
        MsgBox "Colonne '" & mois & " %' non trouvée dans la feuille '" & PERSONNEL_SHEET & "' !", vbCritical
        GoTo CleanUp
    End If
    
    '--- ÉTAPE 3: Créer un "annuaire" du personnel pour une recherche instantanée ---
    lastRowPers = wsPers.Cells(wsPers.Rows.Count, colMatPers).End(xlUp).row
    arrPers = wsPers.Range(wsPers.Cells(1, 1), wsPers.Cells(lastRowPers, colMoisPers)).value
    
    Set personnelDict = CreateObject("Scripting.Dictionary")
    personnelDict.CompareMode = vbTextCompare
    
    For i = 2 To UBound(arrPers, 1) ' Boucle sur le tableau en mémoire (rapide)
        Dim matricule As String
        matricule = Trim(CStr(arrPers(i, colMatPers)))
        If matricule <> "" And Not personnelDict.Exists(matricule) Then
            personnelDict.Add matricule, i ' La clé est le matricule, la valeur est la ligne dans le tableau
        End If
    Next i
    
    '--- ÉTAPE 4: Traiter le planning et calculer les heures ---
    lastRowPlan = wsPlan.Cells(wsPlan.Rows.Count, colNomPlan).End(xlUp).row
    If lastRowPlan < START_ROW Then GoTo CleanUp ' Si pas d'agents
    
    arrPlan = wsPlan.Range(wsPlan.Cells(START_ROW, 1), wsPlan.Cells(lastRowPlan, colHeuresAPrester)).value
    ReDim arrResultats(1 To UBound(arrPlan, 1), 1 To 1)

    For i = 1 To UBound(arrPlan, 1)
        Dim nomPersonne As String, matriculePersonne As String
        nomPersonne = Trim(CStr(arrPlan(i, colNomPlan)))
        matriculePersonne = Trim(CStr(arrPlan(i, colMatPlan)))
        
        If nomPersonne <> "" And matriculePersonne <> "" Then
            If personnelDict.Exists(matriculePersonne) Then
                Dim rowInPersArray As Long
                rowInPersArray = personnelDict(matriculePersonne)
                
                Dim valPourcent As String, pourcentage As Double
                valPourcent = CStr(arrPers(rowInPersArray, colMoisPers))
                valPourcent = Replace(valPourcent, "%", "")
                
                If IsNumeric(valPourcent) Then
                    pourcentage = CDbl(valPourcent) / 100
                Else
                    pourcentage = 1 ' Défaut temps plein
                End If
                
                arrResultats(i, 1) = HeuresAPresterDyn(mois, pourcentage, annee)
            Else
                arrResultats(i, 1) = "Matricule non trouvé" ' Pour le débogage
            End If
        End If
    Next i
    
    '--- ÉTAPE 5: Écrire tous les résultats en une seule fois ---
    wsPlan.Cells(START_ROW, colHeuresAPrester).Resize(UBound(arrResultats, 1), 1).value = arrResultats
    
    MsgBox "Mise à jour des heures à prester pour '" & mois & " " & annee & "' terminée !", vbInformation

CleanUp:
    '--- Nettoyage ---
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Function TrouverColonne(ws As Worksheet, nomHeader As String) As Long
    ' Version améliorée qui cherche dans les 5 premières lignes
    Dim searchRange As Range, foundCell As Range
    Set searchRange = ws.Range("A1:AZ5")
    
    Set foundCell = searchRange.Find(What:=nomHeader, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        TrouverColonne = foundCell.Column
    Else
        TrouverColonne = 0
    End If
End Function


Private Function HeuresAPresterDyn(ByVal nomMois As String, ByVal pct As Double, ByVal annee As Integer) As Double
    ' Calcule le nombre d'heures théoriques pour un mois donné, basé sur les jours ouvrables.
    Dim joursOuvrables As Long
    Dim i As Long, moisNum As Long, nbJoursMois As Long
    
    On Error Resume Next
    moisNum = Month(DateValue("1 " & nomMois & " " & annee))
    If Err.Number <> 0 Then HeuresAPresterDyn = 0: Exit Function
    On Error GoTo 0
    
    nbJoursMois = Day(DateSerial(annee, moisNum + 1, 0))
    
    For i = 1 To nbJoursMois
        ' Compte les jours du Lundi (2) au Vendredi (6) selon le format VBA
        If Weekday(DateSerial(annee, moisNum, i), vbMonday) < 6 Then
            joursOuvrables = joursOuvrables + 1
        End If
    Next i
    
    ' Base de calcul (ex: 7.6h par jour pour un temps plein). Adaptez ce chiffre si nécessaire.
    Const HEURES_JOUR_TEMPS_PLEIN As Double = 7.6
    
    HeuresAPresterDyn = joursOuvrables * HEURES_JOUR_TEMPS_PLEIN * pct
End Function

