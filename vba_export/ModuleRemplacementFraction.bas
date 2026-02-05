Attribute VB_Name = "ModuleRemplacementFraction"
Option Explicit

' =========================================================================
' MODULE: ModuleRemplacementFraction
' BUT:    Analyse et optimisation des plannings de personnel
' AUTEUR: Utilisateur VBA
' DATE:   Mai 2025
' =========================================================================

' =========================================================================
' SECTION 1: CONSTANTES GLOBALES DU MODULE
' =========================================================================

' Indices pour les tableaux de totaux (Matin, Après-midi, Soir)
Const IDX_TOTAL_MATIN As Integer = 0
Const IDX_TOTAL_AM As Integer = 1
Const IDX_TOTAL_SOIR As Integer = 2

' Indices pour les codes de suggestion
Const SUGG_645 As Integer = 0        ' "6:45 15:15"
Const SUGG_7_1530 As Integer = 1     ' "7 15:30"
Const SUGG_7_1130 As Integer = 2     ' "7 11:30"
Const SUGG_7_13 As Integer = 3       ' "7 13"
Const SUGG_8_1630 As Integer = 4     ' "8 16:30"
Const SUGG_C15_GRP As Integer = 5    ' "C 15", "C 15 bis"
Const SUGG_C20_CODE As Integer = 6   ' "C 20"
Const SUGG_C20E_CODE As Integer = 7  ' "C 20 E"
Const SUGG_C19_CODE As Integer = 8   ' "C 19"
Const SUGG_12_30_16_30 As Integer = 9 ' "12:30 16:30"
Const SUGG_NUIT1 As Integer = 10     ' "19:45 6:45"
Const SUGG_NUIT2 As Integer = 11     ' "20 7"

' =========================================================================
' SECTION 2: TYPES DE DONNÉES LOCAUX
' =========================================================================

' Structure pour stocker les normes d'effectif par période
Private Type NormeJour
    Matin As Long
    AM As Long
    Soir As Long
End Type

' =========================================================================
' SECTION 3: FONCTIONS UTILITAIRES
' =========================================================================

' -----------------------------------------------------------------------------
' Fonction: IsJourFerieOuRecup
' But:      Déterminer si un code correspond à un jour férié ou de récupération
' Argument: Le code à vérifier
' Retourne: True si c'est un jour férié ou de récupération, False sinon
' -----------------------------------------------------------------------------
Public Function IsJourFerieOuRecup(code As String) As Boolean
    ' Stub local en attendant modConfigRegles
    Dim codesExclus As Variant
    codesExclus = Array("JF", "F", "RF", "RH", "CA", "CT", "CSS", "MAL", "AM", "AT")
    Dim i As Integer
    For i = LBound(codesExclus) To UBound(codesExclus)
        If StrComp(Trim(code), codesExclus(i), vbTextCompare) = 0 Then
            IsJourFerieOuRecup = True
            Exit Function
        End If
    Next i
    IsJourFerieOuRecup = False
End Function

' -----------------------------------------------------------------------------
' Fonction: GetNomJourSemaine
' But:      Obtenir le nom du jour de la semaine à partir d'une date
' Argument: La date pour laquelle obtenir le nom du jour
' Retourne: Le nom du jour en français
' -----------------------------------------------------------------------------
Public Function GetNomJourSemaine(dateJour As Date) As String
    Dim jourSemaine As Integer
    jourSemaine = Weekday(dateJour)
    
    Select Case jourSemaine
        Case 1: GetNomJourSemaine = "Dimanche"
        Case 2: GetNomJourSemaine = "Lundi"
        Case 3: GetNomJourSemaine = "Mardi"
        Case 4: GetNomJourSemaine = "Mercredi"
        Case 5: GetNomJourSemaine = "Jeudi"
        Case 6: GetNomJourSemaine = "Vendredi"
        Case 7: GetNomJourSemaine = "Samedi"
        Case Else: GetNomJourSemaine = "Jour inconnu"
    End Select
End Function

' -----------------------------------------------------------------------------
' Fonction: GetDateFromColonne
' But:      Obtenir la date correspondant à une colonne dans une feuille de mois
' Arguments:
'   ws:     La feuille de calcul
'   col:    Le numéro de colonne
' Retourne: La date correspondante
' -----------------------------------------------------------------------------
Public Function GetDateFromColonne(ws As Worksheet, col As Integer) As Date
    Dim Mois As Integer, Annee As Integer
    Dim jourMois As Integer
    
    ' Extraire le mois et l'année à partir du nom de la feuille
    Select Case ws.Name
        Case "Janvier": Mois = 1
        Case "Février": Mois = 2
        Case "Mars": Mois = 3
        Case "Avril": Mois = 4
        Case "Mai": Mois = 5
        Case "Juin": Mois = 6
        Case "Juillet": Mois = 7
        Case "Août": Mois = 8
        Case "Septembre": Mois = 9
        Case "Octobre": Mois = 10
        Case "Novembre": Mois = 11
        Case "Décembre": Mois = 12
        Case Else: Mois = 1 ' Par défaut
    End Select
    
    ' Récupérer l'année depuis la cellule appropriée
    On Error Resume Next
    Annee = ws.Range("B1").value
    If Err.Number <> 0 Or Annee = 0 Then
        Annee = Year(Now) ' Utiliser l'année courante si non trouvée
    End If
    On Error GoTo 0
    
    ' Le jour du mois est égal au numéro de colonne
    jourMois = col
    
    ' Créer et retourner la date
    GetDateFromColonne = DateSerial(Annee, Mois, jourMois)
End Function

' -----------------------------------------------------------------------------
' Fonction: ObtenirNormesJour (locale)
' But:      Obtenir les normes d'effectif pour un type de jour
' Arguments:
'   jourSemaine: Numéro du jour (1=Dimanche, 2=Lundi, etc.)
'   estFerie:    True si c'est un jour férié
' Retourne: Structure NormeJour avec les effectifs cibles
' -----------------------------------------------------------------------------
Private Function ObtenirNormesJour(jourSemaine As Integer, estFerie As Boolean) As NormeJour
    Dim normes As NormeJour
    
    ' Valeurs par défaut pour un jour férié
    If estFerie Then
        normes.Matin = 2
        normes.AM = 2
        normes.Soir = 1
        ObtenirNormesJour = normes
        Exit Function
    End If
    
    ' Valeurs selon le jour de la semaine
    Select Case jourSemaine
        Case 1 ' Dimanche
            normes.Matin = 3
            normes.AM = 3
            normes.Soir = 2
        Case 7 ' Samedi
            normes.Matin = 4
            normes.AM = 4
            normes.Soir = 2
        Case Else ' Lundi à Vendredi
            normes.Matin = 6
            normes.AM = 6
            normes.Soir = 3
    End Select
    
    ObtenirNormesJour = normes
End Function

' =========================================================================
' SECTION 4: TRAITEMENT DES FEUILLES DE MOIS
' =========================================================================

' -----------------------------------------------------------------------------
' Procedure: TraiterUneFeuilleDeMois
' But:       Analyser et optimiser une feuille de mois pour les remplacements
' Arguments:
'   ws:      La feuille de calcul à traiter
'   debug:   Si True, affiche des informations de débogage
' -----------------------------------------------------------------------------
Public Sub TraiterUneFeuilleDeMois(ws As Worksheet, Optional debug As Boolean = False)
    ' Déléguer à la version qui utilise les règles de configuration
    TraiterUneFeuilleDeMois_ParRegles ws, debug
End Sub

' -----------------------------------------------------------------------------
' Procedure: TraiterUneFeuilleDeMois_ParRegles
' But:       Version améliorée qui utilise les règles définies
' Arguments:
'   ws:      La feuille de calcul à traiter
'   debug:   Si True, affiche des informations de débogage
' -----------------------------------------------------------------------------
' -----------------------------------------------------------------------------
' Procedure: TraiterUneFeuilleDeMois
' But:       Analyser et optimiser une feuille de mois pour les remplacements
' Arguments:
'   ws:      La feuille de calcul à traiter
'   debug:   Si True, affiche des informations de débogage
' -----------------------------------------------------------------------------
Public Sub TraiterUneFeuilleDeMois(ws As Worksheet, Optional debug As Boolean = False)
    Dim nbLignes As Long, nbJours As Integer
    Dim row As Long, col As Integer
    Dim cellule As Range
    Dim code As String
    
    ' Tableaux pour stocker les totaux par période et par jour
    Dim totauxMASArr() As Variant
    
    ' Tableaux pour stocker les cibles d'effectifs par jour
    Dim arrTargetMatin() As Long
    Dim arrTargetPM() As Long
    Dim arrTargetSoir() As Long
    
    ' Variables pour les calculs de manque
    Dim actualMatin As Long, actualPM As Long, actualSoir As Long
    Dim targetMatin As Long, targetPM As Long, targetSoir As Long
    Dim manqueMatin As Long, manquePM As Long, manqueSoir As Long
    
    ' Déterminer le nombre de jours dans le mois
    nbJours = 31
    
    ' Déterminer le nombre de lignes à analyser
    nbLignes = ws.UsedRange.Rows.count
    If nbLignes < 10 Then Exit Sub
    
    ' Initialiser les tableaux de totaux
    ReDim totauxMASArr(0 To 2, 1 To nbJours)
    
    ' Initialiser les tableaux de cibles
    ReDim arrTargetMatin(1 To nbJours)
    ReDim arrTargetPM(1 To nbJours)
    ReDim arrTargetSoir(1 To nbJours)
    
    ' Calculer les totaux pour chaque jour et période
    For row = 1 To nbLignes
        For col = 1 To nbJours
            Set cellule = ws.Cells(row, col + 1)
            
            If Not IsEmpty(cellule) Then
                code = Trim(cellule.value)
                
                If Len(code) > 0 And Not IsNumeric(code) Then
                    If EstCodeMatin(code) Then
                        totauxMASArr(IDX_TOTAL_MATIN, col) = totauxMASArr(IDX_TOTAL_MATIN, col) + 1
                    End If
                    
                    If EstCodeApresMidi(code) Then
                        totauxMASArr(IDX_TOTAL_AM, col) = totauxMASArr(IDX_TOTAL_AM, col) + 1
                    End If
                    
                    If EstCodeSoir(code) Then
                        totauxMASArr(IDX_TOTAL_SOIR, col) = totauxMASArr(IDX_TOTAL_SOIR, col) + 1
                    End If
                End If
            End If
        Next col
    Next row
    
    ' Déterminer les cibles pour chaque jour
    For col = 1 To nbJours
        Dim dateJour As Date
        Dim jourSemaine As Integer
        Dim estFerie As Boolean
        Dim codeJourFerie As String
        Dim normes As NormeJour
        
        dateJour = GetDateFromColonne(ws, col)
        jourSemaine = Weekday(dateJour)
        codeJourFerie = ObtenirCodeJourFerie(dateJour)
        estFerie = (codeJourFerie <> "")
        normes = ObtenirNormesJour(jourSemaine, estFerie)
        
        arrTargetMatin(col) = normes.Matin
        arrTargetPM(col) = normes.AM
        arrTargetSoir(col) = normes.Soir
    Next col
    
    ' Analyser les manques et suggérer des remplacements
    For col = 1 To nbJours
        If IsNumeric(totauxMASArr(IDX_TOTAL_MATIN, col)) Then actualMatin = CLng(totauxMASArr(IDX_TOTAL_MATIN, col)) Else actualMatin = 0
        If IsNumeric(totauxMASArr(IDX_TOTAL_AM, col)) Then actualPM = CLng(totauxMASArr(IDX_TOTAL_AM, col)) Else actualPM = 0
        If IsNumeric(totauxMASArr(IDX_TOTAL_SOIR, col)) Then actualSoir = CLng(totauxMASArr(IDX_TOTAL_SOIR, col)) Else actualSoir = 0
        
        targetMatin = arrTargetMatin(col)
        targetPM = arrTargetPM(col)
        targetSoir = arrTargetSoir(col)
        
        manqueMatin = targetMatin - actualMatin
        manquePM = targetPM - actualPM
        manqueSoir = targetSoir - actualSoir
        
        If manqueMatin > 0 Or manquePM > 0 Or manqueSoir > 0 Then
            AnalyseEtRemplacementPlanningUltraOptimise ws, col, manqueMatin, manquePM, manqueSoir, debug
        End If
    Next col
    
    If debug Then
        Dim msg As String
        msg = "Traitement terminé pour la feuille " & ws.Name & vbCrLf
        msg = msg & "Nombre de jours analysés: " & nbJours & vbCrLf
        MsgBox msg, vbInformation, "Résumé du traitement"
    End If
End Sub

' -----------------------------------------------------------------------------
' Fonction: ObtenirCodeJourFerie
' But:      Obtenir le code de jour férié pour une date donnée
' Argument: La date à vérifier
' Retourne: Le code du jour férié ou chaîne vide si ce n'est pas un jour férié
' -----------------------------------------------------------------------------
Private Function ObtenirCodeJourFerie(dateJour As Date) As String
    Dim jour As Integer, Mois As Integer
    
    jour = Day(dateJour)
    Mois = Month(dateJour)
    
    ' Jours fériés fixes
    If jour = 1 And Mois = 1 Then
        ObtenirCodeJourFerie = "F 1-1"    ' Jour de l'an
        Exit Function
    End If
    If jour = 1 And Mois = 5 Then
        ObtenirCodeJourFerie = "F 1-5"    ' Fête du travail
        Exit Function
    End If
    If jour = 8 And Mois = 5 Then
        ObtenirCodeJourFerie = "F 8-5"    ' Victoire 1945
        Exit Function
    End If
    If jour = 14 And Mois = 7 Then
        ObtenirCodeJourFerie = "F 14-7"  ' Fête nationale
        Exit Function
    End If
    If jour = 15 And Mois = 8 Then
        ObtenirCodeJourFerie = "F 15-8"  ' Assomption
        Exit Function
    End If
    If jour = 1 And Mois = 11 Then
        ObtenirCodeJourFerie = "F 1-11"  ' Toussaint
        Exit Function
    End If
    If jour = 11 And Mois = 11 Then
        ObtenirCodeJourFerie = "F 11-11" ' Armistice
        Exit Function
    End If
    If jour = 25 And Mois = 12 Then
        ObtenirCodeJourFerie = "F 25-12" ' Noël
        Exit Function
    End If
    
    ' Pas un jour férié
    ObtenirCodeJourFerie = ""
End Function

' -----------------------------------------------------------------------------
' Fonction: EstCodeMatin
' But:      Déterminer si un code couvre la période du matin (6h-14h)
' Argument: Le code à vérifier
' Retourne: True si le code couvre la période du matin, False sinon
' -----------------------------------------------------------------------------
Private Function EstCodeMatin(code As String) As Boolean
    ' Liste des codes qui couvrent la période du matin
    Dim codesMatin As Variant
    codesMatin = Array("6:45 15:15", "7 15:30", "7 11:30", "7 13", "8 16:30", "6 14", "7 14")
    
    ' Vérifier si le code est dans la liste
    Dim i As Integer
    For i = LBound(codesMatin) To UBound(codesMatin)
        If StrComp(code, codesMatin(i), vbTextCompare) = 0 Then
            EstCodeMatin = True
            Exit Function
        End If
    Next i
    
    ' Vérifier les cas spéciaux
    If Left(code, 1) = "6" Or Left(code, 1) = "7" Or Left(code, 1) = "8" Then
        EstCodeMatin = True
        Exit Function
    End If
    
    ' Par défaut, le code ne couvre pas la période du matin
    EstCodeMatin = False
End Function

' -----------------------------------------------------------------------------
' Fonction: EstCodeApresMidi
' But:      Déterminer si un code couvre la période de l'après-midi (14h-20h)
' Argument: Le code à vérifier
' Retourne: True si le code couvre la période de l'après-midi, False sinon
' -----------------------------------------------------------------------------
Private Function EstCodeApresMidi(code As String) As Boolean
    ' Liste des codes qui couvrent la période de l'après-midi
    Dim codesAM As Variant
    codesAM = Array("6:45 15:15", "7 15:30", "8 16:30", "12:30 16:30", "14 22", "13 21", "C 15", "C 15 bis")
    
    ' Vérifier si le code est dans la liste
    Dim i As Integer
    For i = LBound(codesAM) To UBound(codesAM)
        If StrComp(code, codesAM(i), vbTextCompare) = 0 Then
            EstCodeApresMidi = True
            Exit Function
        End If
    Next i
    
    ' Vérifier les cas spéciaux
    If Left(code, 1) = "C" And (InStr(code, "15") > 0 Or InStr(code, "14") > 0) Then
        EstCodeApresMidi = True
        Exit Function
    End If
    
    ' Par défaut, le code ne couvre pas la période de l'après-midi
    EstCodeApresMidi = False
End Function

' -----------------------------------------------------------------------------
' Fonction: EstCodeSoir
' But:      Déterminer si un code couvre la période du soir (20h-6h)
' Argument: Le code à vérifier
' Retourne: True si le code couvre la période du soir, False sinon
' -----------------------------------------------------------------------------
Private Function EstCodeSoir(code As String) As Boolean
    ' Liste des codes qui couvrent la période du soir
    Dim codesSoir As Variant
    codesSoir = Array("19:45 6:45", "20 7", "C 19", "C 20", "C 20 E", "22 6", "21 5")
    
    ' Vérifier si le code est dans la liste
    Dim i As Integer
    For i = LBound(codesSoir) To UBound(codesSoir)
        If StrComp(code, codesSoir(i), vbTextCompare) = 0 Then
            EstCodeSoir = True
            Exit Function
        End If
    Next i
    
    ' Vérifier les cas spéciaux
    If Left(code, 1) = "C" And (InStr(code, "19") > 0 Or InStr(code, "20") > 0) Then
        EstCodeSoir = True
        Exit Function
    End If
    
    ' Par défaut, le code ne couvre pas la période du soir
    EstCodeSoir = False
End Function

' =========================================================================
' SECTION 5: ANALYSE ET REMPLACEMENT
' =========================================================================

' -----------------------------------------------------------------------------
' Procedure: AnalyseEtRemplacementPlanningUltraOptimise
' But:       Analyser les manques et suggérer des remplacements optimisés
' Arguments:
'   ws:          La feuille de calcul à traiter
'   colJour:     La colonne correspondant au jour à analyser
'   manqueMatin: Le nombre de personnes manquantes le matin
'   manquePM:    Le nombre de personnes manquantes l'après-midi
'   manqueSoir:  Le nombre de personnes manquantes le soir
'   debug:       Si True, affiche des informations de débogage
' -----------------------------------------------------------------------------
Public Sub AnalyseEtRemplacementPlanningUltraOptimise(ws As Worksheet, colJour As Integer, _
                                                    manqueMatin As Long, manquePM As Long, manqueSoir As Long, _
                                                    Optional debug As Boolean = False)
    Dim row As Long, lastRow As Long
    Dim dateJour As Date
    Dim jourSemaine As Integer
    Dim nomJour As String
    Dim estFerie As Boolean
    Dim codeJourFerie As String
    Dim i As Integer
    Dim suggestionFaite As Boolean
    Dim nbSuggestions As Integer
    Dim lignesLibres() As Long
    Dim nbLignesLibres As Integer
    
    ' Obtenir la date correspondant à cette colonne
    dateJour = GetDateFromColonne(ws, colJour)
    
    ' Déterminer le jour de la semaine
    jourSemaine = Weekday(dateJour)
    nomJour = GetNomJourSemaine(dateJour)
    
    ' Vérifier si c'est un jour férié
    codeJourFerie = ObtenirCodeJourFerie(dateJour)
    estFerie = (codeJourFerie <> "")
    
    ' Trouver la dernière ligne utilisée dans la feuille
    lastRow = ws.UsedRange.Rows.count
    
    ' Initialiser le tableau pour stocker les lignes libres
    ReDim lignesLibres(1 To lastRow)
    nbLignesLibres = 0
    
    ' Identifier les lignes libres pour ce jour
    For row = 1 To lastRow
        ' Vérifier si la cellule est vide
        If IsEmpty(ws.Cells(row, colJour + 1)) Then
            ' Vérifier si c'est une ligne de personnel
            If Not IsEmpty(ws.Cells(row, 1)) And Len(Trim(ws.Cells(row, 1).value)) > 0 Then
                ' Ajouter cette ligne à notre tableau de lignes libres
                nbLignesLibres = nbLignesLibres + 1
                lignesLibres(nbLignesLibres) = row
            End If
        End If
    Next row
    
    ' Si aucune ligne libre, sortir
    If nbLignesLibres = 0 Then
        If debug Then MsgBox "Aucune ligne libre trouvée pour le " & nomJour & " " & Day(dateJour) & "/" & Month(dateJour), vbInformation
        Exit Sub
    End If
    
    ' Initialiser le compteur de suggestions
    nbSuggestions = 0
    suggestionFaite = False
    
    ' Traiter les manques de nuit en priorité
    If manqueSoir > 0 Then
        suggestionFaite = SuggererRemplacementsNuit(ws, colJour, manqueSoir, lignesLibres, nbLignesLibres, nbSuggestions, debug)
    End If
    
    ' Traiter les manques de matin et après-midi
    If manqueMatin > 0 Or manquePM > 0 Then
        suggestionFaite = SuggererRemplacementsJour(ws, colJour, manqueMatin, manquePM, lignesLibres, nbLignesLibres, nbSuggestions, debug) Or suggestionFaite
    End If
    
    ' Afficher un résumé si en mode debug
    If debug Then
        Dim msg As String
        msg = "Analyse pour le " & nomJour & " " & Day(dateJour) & "/" & Month(dateJour) & vbCrLf
        msg = msg & "Manque Matin: " & manqueMatin & vbCrLf
        msg = msg & "Manque PM: " & manquePM & vbCrLf
        msg = msg & "Manque Soir: " & manqueSoir & vbCrLf
        msg = msg & "Suggestions faites: " & nbSuggestions & vbCrLf
        MsgBox msg, vbInformation, "Résumé des suggestions"
    End If
End Sub

' -----------------------------------------------------------------------------
' Fonction: SuggererRemplacementsNuit
' But:      Suggérer des remplacements pour les manques de nuit
' Arguments:
'   ws:           La feuille de calcul à traiter
'   colJour:      La colonne correspondant au jour à analyser
'   manqueSoir:   Le nombre de personnes manquantes le soir
'   lignesLibres: Tableau des lignes disponibles pour les remplacements
'   nbLignesLibres: Nombre de lignes disponibles
'   nbSuggestions: Nombre de suggestions faites (modifié par référence)
'   debug:        Si True, affiche des informations de débogage
' Retourne: True si au moins une suggestion a été faite, False sinon
' -----------------------------------------------------------------------------
Private Function SuggererRemplacementsNuit(ws As Worksheet, colJour As Integer, _
                                         manqueSoir As Long, lignesLibres() As Long, _
                                         ByRef nbLignesLibres As Integer, ByRef nbSuggestions As Integer, _
                                         Optional debug As Boolean = False) As Boolean
    Dim i As Integer
    Dim suggestionsFaites As Integer
    Dim codeNuit As String
    
    ' Initialiser le compteur de suggestions
    suggestionsFaites = 0
    
    ' Codes de nuit à suggérer
    Dim codesNuit As Variant
    codesNuit = Array("19:45 6:45", "20 7", "C 19", "C 20", "C 20 E")
    
    ' Suggérer des remplacements tant qu'il y a des manques et des lignes libres
    For i = 1 To manqueSoir
        If i <= nbLignesLibres Then
            ' Choisir un code de nuit (alternance entre les différents codes)
            codeNuit = codesNuit((i - 1) Mod (UBound(codesNuit) + 1))
            
            ' Appliquer le code à la ligne libre
            ws.Cells(lignesLibres(i), colJour + 1).value = codeNuit
            
            ' Mettre à jour les compteurs
            suggestionsFaites = suggestionsFaites + 1
        Else
            ' Plus de lignes libres disponibles
            Exit For
        End If
    Next i
    
    ' Mettre à jour le nombre de lignes libres restantes
    If suggestionsFaites > 0 Then
        ' Décaler les lignes libres restantes
        For i = 1 To nbLignesLibres - suggestionsFaites
            lignesLibres(i) = lignesLibres(i + suggestionsFaites)
        Next i
        
        ' Mettre à jour le nombre de lignes libres
        nbLignesLibres = nbLignesLibres - suggestionsFaites
        
        ' Mettre à jour le nombre total de suggestions
        nbSuggestions = nbSuggestions + suggestionsFaites
    End If
    
    ' Retourner True si au moins une suggestion a été faite
    SuggererRemplacementsNuit = (suggestionsFaites > 0)
End Function

' -----------------------------------------------------------------------------
' Fonction: SuggererRemplacementsJour
' But:      Suggérer des remplacements pour les manques de jour (matin et après-midi)
' Arguments:
'   ws:           La feuille de calcul à traiter
'   colJour:      La colonne correspondant au jour à analyser
'   manqueMatin:  Le nombre de personnes manquantes le matin
'   manquePM:     Le nombre de personnes manquantes l'après-midi
'   lignesLibres: Tableau des lignes disponibles pour les remplacements
'   nbLignesLibres: Nombre de lignes disponibles
'   nbSuggestions: Nombre de suggestions faites (modifié par référence)
'   debug:        Si True, affiche des informations de débogage
' Retourne: True si au moins une suggestion a été faite, False sinon
' -----------------------------------------------------------------------------
Private Function SuggererRemplacementsJour(ws As Worksheet, colJour As Integer, _
                                         manqueMatin As Long, manquePM As Long, _
                                         lignesLibres() As Long, ByRef nbLignesLibres As Integer, _
                                         ByRef nbSuggestions As Integer, Optional debug As Boolean = False) As Boolean
    Dim i As Integer
    Dim suggestionsFaites As Integer
    Dim code As String
    Dim manqueMatinRestant As Long, manquePMRestant As Long
    
    ' Initialiser les variables
    suggestionsFaites = 0
    manqueMatinRestant = manqueMatin
    manquePMRestant = manquePM
    
    ' Codes pour le matin et l'après-midi
    Dim codesMatin As Variant, codesAM As Variant, codesJournee As Variant
    codesMatin = Array("7 11:30", "7 13")
    codesAM = Array("12:30 16:30", "C 15", "C 15 bis")
    codesJournee = Array("6:45 15:15", "7 15:30", "8 16:30")
    
    ' Traiter d'abord les manques sur toute la journée
    If manqueMatinRestant > 0 And manquePMRestant > 0 Then
        Dim nbJournee As Integer
        nbJournee = Application.WorksheetFunction.Min(manqueMatinRestant, manquePMRestant, nbLignesLibres)
        
        For i = 1 To nbJournee
            ' Choisir un code de journée
            code = codesJournee((i - 1) Mod (UBound(codesJournee) + 1))
            
            ' Appliquer le code à la ligne libre
            ws.Cells(lignesLibres(i), colJour + 1).value = code
            
            ' Mettre à jour les compteurs
            suggestionsFaites = suggestionsFaites + 1
            manqueMatinRestant = manqueMatinRestant - 1
            manquePMRestant = manquePMRestant - 1
        Next i
        
        ' Mettre à jour le nombre de lignes libres restantes
        If nbJournee > 0 Then
            ' Décaler les lignes libres restantes
            For i = 1 To nbLignesLibres - nbJournee
                lignesLibres(i) = lignesLibres(i + nbJournee)
            Next i
            
            ' Mettre à jour le nombre de lignes libres
            nbLignesLibres = nbLignesLibres - nbJournee
        End If
    End If
    
    ' Traiter ensuite les manques restants au matin
    If manqueMatinRestant > 0 And nbLignesLibres > 0 Then
        Dim nbMatin As Integer
        nbMatin = Application.WorksheetFunction.Min(manqueMatinRestant, nbLignesLibres)
        
        For i = 1 To nbMatin
            ' Choisir un code de matin
            code = codesMatin((i - 1) Mod (UBound(codesMatin) + 1))
            
            ' Appliquer le code à la ligne libre
            ws.Cells(lignesLibres(i), colJour + 1).value = code
            
            ' Mettre à jour les compteurs
            suggestionsFaites = suggestionsFaites + 1
            manqueMatinRestant = manqueMatinRestant - 1
        Next i
        
        ' Mettre à jour le nombre de lignes libres restantes
        If nbMatin > 0 Then
            ' Décaler les lignes libres restantes
            For i = 1 To nbLignesLibres - nbMatin
                lignesLibres(i) = lignesLibres(i + nbMatin)
            Next i
            
            ' Mettre à jour le nombre de lignes libres
            nbLignesLibres = nbLignesLibres - nbMatin
        End If
    End If
    
    ' Traiter enfin les manques restants à l'après-midi
    If manquePMRestant > 0 And nbLignesLibres > 0 Then
        Dim nbAM As Integer
        nbAM = Application.WorksheetFunction.Min(manquePMRestant, nbLignesLibres)
        
        For i = 1 To nbAM
            ' Choisir un code d'après-midi
            code = codesAM((i - 1) Mod (UBound(codesAM) + 1))
            
            ' Appliquer le code à la ligne libre
            ws.Cells(lignesLibres(i), colJour + 1).value = code
            
            ' Mettre à jour les compteurs
            suggestionsFaites = suggestionsFaites + 1
            manquePMRestant = manquePMRestant - 1
        Next i
        
        ' Mettre à jour le nombre de lignes libres restantes
        If nbAM > 0 Then
            ' Décaler les lignes libres restantes
            For i = 1 To nbLignesLibres - nbAM
                lignesLibres(i) = lignesLibres(i + nbAM)
            Next i
            
            ' Mettre à jour le nombre de lignes libres
            nbLignesLibres = nbLignesLibres - nbAM
        End If
    End If
    
    ' Mettre à jour le nombre total de suggestions
    nbSuggestions = nbSuggestions + suggestionsFaites
    
    ' Retourner True si au moins une suggestion a été faite
    SuggererRemplacementsJour = (suggestionsFaites > 0)
End Function

