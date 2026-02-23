Attribute VB_Name = "Module_MAJ_HeuresAPrester"
' Version: v2.0 - 2026-02-23
' Changements v2 vs v1 :
'   1. HeuresAPresterDyn() appelle Module_HeuresTravaillees.HeuresTheoriquesMois()
'      au lieu du calcul joursOuvrables * 7.6 * pct (qui ignorait les feries belges)
'   2. HeuresStdJour lu depuis la feuille Personnel (colonne D) au lieu de hardcoder 7.6
'   3. Fallback automatique sur l'ancien calcul si Module_HeuresTravaillees est absent
Option Explicit

'===================================================================================
' MODULE :      Module_MAJ_HeuresAPrester (Version 2.0 -- avec feries belges)
' DESCRIPTION : Met a jour les "Heures a prester" de maniere optimisee en se
'               basant sur la feuille "Personnel". Utilise Module_HeuresTravaillees
'               pour le calcul des heures theoriques (jours feries belges inclus).
'               Fallback local si Module_HeuresTravaillees non disponible.
'
' DEPENDANCES :
'   - Module_HeuresTravaillees : HeuresTheoriquesMois() (calcul avec feries)
'   - Module_Planning_Core     : BuildFeriesBE() (via Module_HeuresTravaillees)
'   - Feuille "Personnel"      : Matricule (col A), HeuresStdJour (col D), % mensuel
'   - Feuille "Accueil"        : Annee en F22
'===================================================================================

'--- Constantes ---
Private Const PERSONNEL_SHEET As String = "Personnel"
Private Const ACCUEIL_SHEET As String = "Accueil"
Private Const YEAR_CELL As String = "F22"
Private Const START_ROW As Long = 6
Private Const MONTH_SHEETS As String = "Janv,Fev,Mars,Avril,Mai,Juin,Juil,Aout,Sept,Oct,Nov,Dec"
Private Const COL_HEURES_STD_JOUR As Long = 4   ' Colonne D dans Personnel = HeuresStdJour par agent

Public Sub MAJ_HeuresAPrester()
    '--- Declaration des variables ---
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

    '--- ETAPE 1: Preparation et validation robuste des feuilles ---
    ' Cette boucle est plus tolerante aux espaces dans les noms d'onglets
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
               "Veuillez verifier les noms des onglets.", vbCritical
        GoTo CleanUp
    End If

    annee = wsAccueil.Range(YEAR_CELL).Value
    If Not IsNumeric(annee) Or annee < 2000 Then
        MsgBox "Annee du planning introuvable ou incorrecte dans " & ACCUEIL_SHEET & "!" & YEAR_CELL, vbCritical
        GoTo CleanUp
    End If

    '--- ETAPE 2: Trouver les colonnes necessaires (de maniere robuste) ---
    Set wsPlan = ActiveSheet
    colNomPlan = TrouverColonne(wsPlan, "Nom")
    colMatPlan = TrouverColonne(wsPlan, "Matricule")
    colHeuresAPrester = TrouverColonne(wsPlan, "Heures " & ChrW(224) & " prester")
    If colHeuresAPrester = 0 Then
        colHeuresAPrester = TrouverColonne(wsPlan, "Heures a prester")
    End If
    If colHeuresAPrester = 0 Then
        colHeuresAPrester = TrouverColonne(wsPlan, "Heures")
    End If
    colMatPers = 1 ' On suppose que le matricule est toujours en colonne A de "Personnel"

    If colNomPlan = 0 Or colMatPlan = 0 Or colHeuresAPrester = 0 Then
        MsgBox "Une colonne cle ('Nom', 'Matricule' ou 'Heures a prester') n'a pas ete trouvee sur le planning !", vbCritical
        GoTo CleanUp
    End If

    ' Trouver la colonne du mois dans "Personnel" (ex: "Janv %")
    mois = wsPlan.Name
    For i = 1 To wsPers.Cells(1, wsPers.Columns.Count).End(xlToLeft).Column
        If InStr(1, wsPers.Cells(1, i).Value, mois, vbTextCompare) > 0 And InStr(1, wsPers.Cells(1, i).Value, "%") > 0 Then
            colMoisPers = i
            Exit For
        End If
    Next i

    If colMoisPers = 0 Then
        MsgBox "Colonne '" & mois & " %' non trouvee dans la feuille '" & PERSONNEL_SHEET & "' !", vbCritical
        GoTo CleanUp
    End If

    '--- ETAPE 3: Creer un "annuaire" du personnel pour une recherche instantanee ---
    ' On etend la plage lue pour inclure la colonne HeuresStdJour (col D = 4)
    Dim maxColPers As Long
    maxColPers = colMoisPers
    If COL_HEURES_STD_JOUR > maxColPers Then maxColPers = COL_HEURES_STD_JOUR

    lastRowPers = wsPers.Cells(wsPers.Rows.Count, colMatPers).End(xlUp).Row
    arrPers = wsPers.Range(wsPers.Cells(1, 1), wsPers.Cells(lastRowPers, maxColPers)).Value

    Set personnelDict = CreateObject("Scripting.Dictionary")
    personnelDict.CompareMode = vbTextCompare

    For i = 2 To UBound(arrPers, 1) ' Boucle sur le tableau en memoire (rapide)
        Dim matricule As String
        matricule = Trim(CStr(arrPers(i, colMatPers)))
        If matricule <> "" And Not personnelDict.Exists(matricule) Then
            personnelDict.Add matricule, i ' La cle est le matricule, la valeur est la ligne dans le tableau
        End If
    Next i

    '--- ETAPE 4: Convertir le nom de mois en numero (1-12) ---
    Dim moisNum As Long
    moisNum = GetMoisNum(mois)
    If moisNum = 0 Then
        MsgBox "Nom d'onglet '" & mois & "' non reconnu comme un mois valide.", vbCritical
        GoTo CleanUp
    End If

    '--- ETAPE 5: Traiter le planning et calculer les heures ---
    lastRowPlan = wsPlan.Cells(wsPlan.Rows.Count, colNomPlan).End(xlUp).Row
    If lastRowPlan < START_ROW Then GoTo CleanUp ' Si pas d'agents

    arrPlan = wsPlan.Range(wsPlan.Cells(START_ROW, 1), wsPlan.Cells(lastRowPlan, colHeuresAPrester)).Value
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
                    pourcentage = 1 ' Defaut temps plein
                End If

                ' Lire heuresStdJour depuis Personnel (colonne D) pour cet agent
                Dim heuresStdAgent As Double
                heuresStdAgent = 7.6 ' Valeur par defaut secteur hospitalier belge
                If COL_HEURES_STD_JOUR <= UBound(arrPers, 2) Then
                    Dim valStd As Variant
                    valStd = arrPers(rowInPersArray, COL_HEURES_STD_JOUR)
                    If IsNumeric(valStd) Then
                        If CDbl(valStd) > 0 And CDbl(valStd) <= 12 Then
                            heuresStdAgent = CDbl(valStd)
                        End If
                    End If
                End If

                arrResultats(i, 1) = HeuresAPresterDyn(mois, pourcentage, annee, moisNum, heuresStdAgent)
            Else
                arrResultats(i, 1) = "Matricule non trouve" ' Pour le debogage
            End If
        End If
    Next i

    '--- ETAPE 6: Ecrire tous les resultats en une seule fois ---
    wsPlan.Cells(START_ROW, colHeuresAPrester).Resize(UBound(arrResultats, 1), 1).Value = arrResultats

    MsgBox "Mise a jour des heures a prester pour '" & mois & " " & annee & "' terminee !" & vbCrLf & _
           "(v2 : jours feries belges pris en compte)", vbInformation

CleanUp:
    '--- Nettoyage ---
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Function TrouverColonne(ws As Worksheet, nomHeader As String) As Long
    ' Version amelioree qui cherche dans les 5 premieres lignes
    Dim searchRange As Range, foundCell As Range
    Set searchRange = ws.Range("A1:AZ5")

    Set foundCell = searchRange.Find(What:=nomHeader, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

    If Not foundCell Is Nothing Then
        TrouverColonne = foundCell.Column
    Else
        TrouverColonne = 0
    End If
End Function


Private Function HeuresAPresterDyn(ByVal nomMois As String, _
                                    ByVal pct As Double, _
                                    ByVal annee As Integer, _
                                    ByVal moisNum As Long, _
                                    ByVal heuresStdJour As Double) As Double
    ' v2: Appelle Module_HeuresTravaillees.HeuresTheoriquesMois() pour un calcul
    ' precis incluant les jours feries belges (via BuildFeriesBE).
    '
    ' HeuresTheoriquesMois(annee, moisNum, pctTemps, heuresStdJour) retourne:
    '   joursOuvrables (hors feries) * heuresStdJour
    ' Le pctTemps n'est utilise que pour validation dans HeuresTheoriquesMois.
    ' Donc on passe heuresStdJour * pct pour obtenir le bon resultat pour les
    ' temps partiels (ex: 7.6 * 0.75 = 5.7h/jour pour un 75%).

    On Error GoTo Fallback

    ' --- Nouveau calcul avec feries belges ---
    HeuresAPresterDyn = Module_HeuresTravaillees.HeuresTheoriquesMois( _
        CLng(annee), moisNum, 1#, heuresStdJour * pct)
    Exit Function

Fallback:
    ' --- Ancien calcul sans feries (compatibilite si Module_HeuresTravaillees absent) ---
    On Error GoTo 0
    Dim joursOuvrables As Long
    Dim j As Long, nbJoursMois As Long

    If moisNum < 1 Or moisNum > 12 Then
        ' Essayer de deduire moisNum depuis le nom du mois
        On Error Resume Next
        moisNum = Month(DateValue("1 " & nomMois & " " & annee))
        If Err.Number <> 0 Then
            HeuresAPresterDyn = 0
            Exit Function
        End If
        On Error GoTo 0
    End If

    nbJoursMois = Day(DateSerial(annee, moisNum + 1, 0))
    joursOuvrables = 0

    For j = 1 To nbJoursMois
        ' Compte les jours du Lundi (1) au Vendredi (5) avec vbMonday
        If Weekday(DateSerial(annee, moisNum, j), vbMonday) < 6 Then
            joursOuvrables = joursOuvrables + 1
        End If
    Next j

    ' Ancien calcul: jours ouvrables * heuresStdJour * pourcentage (sans feries)
    HeuresAPresterDyn = joursOuvrables * heuresStdJour * pct
End Function


Private Function GetMoisNum(ByVal nomMois As String) As Long
    ' Convertit un nom d'onglet mois en numero 1-12
    ' Supporte les noms d'onglets du planning: Janv, Fev, Mars, Avril, etc.
    Dim mSheets() As String
    Dim i As Long
    mSheets = Split(MONTH_SHEETS, ",")
    For i = 0 To 11
        If UCase(Trim(mSheets(i))) = UCase(Trim(nomMois)) Then
            GetMoisNum = i + 1
            Exit Function
        End If
    Next i
    ' Fallback: essayer VBA DateValue pour les noms complets (Janvier, Fevrier...)
    On Error Resume Next
    GetMoisNum = Month(DateValue("1 " & nomMois & " 2026"))
    If Err.Number <> 0 Then GetMoisNum = 0
    On Error GoTo 0
End Function
