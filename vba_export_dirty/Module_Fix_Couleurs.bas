Attribute VB_Name = "Module_Fix_Couleurs"
Option Explicit

' ========================================================================================
' SCRIPT DE CORRECTION AUTOMATIQUE DES COULEURS
' Detecte la couleur exacte du "Bleu Bains" sur la cellule de Mamadou (ou Edelyne)
' et met a jour la configuration.
' ========================================================================================

Sub Corriger_Couleur_Bleu_Clair()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim nomPersonne As String
    nomPersonne = "Diallo_Mamadou" ' Nom cible pour la detection
    
    Dim cPersonne As Range
    Dim rFound As Range
    Dim colJour As Long
    Dim couleurDetectee As Long
    Dim found As Boolean
    
    ' 1. Trouver la personne (Mamadou)
    ' On cherche dans la colonne B (Nom) ou A
    Set rFound = ws.Columns("A:B").Find(nomPersonne, LookIn:=xlValues, LookAt:=xlPart)
    
    If rFound Is Nothing Then
        ' Essayer avec Edelyne
        nomPersonne = "Dela Vega_Edelyn"
        Set rFound = ws.Columns("A:B").Find(nomPersonne, LookIn:=xlValues, LookAt:=xlPart)
    End If
    
    If rFound Is Nothing Then
        MsgBox "Impossible de trouver 'Diallo_Mamadou' ou 'Dela Vega_Edelyn' sur cette feuille.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Trouver la colonne du Mercredi 4 (valeur 4 en ligne des jours)
    ' La ligne des jours est generalement la 4 (PLN_Row_DayNumbers)
    Dim ligneJours As Long: ligneJours = 4
    
    found = False
    For colJour = 3 To 35 ' Colonnes C a AI
        If IsNumeric(ws.Cells(ligneJours, colJour).value) Then
            If ws.Cells(ligneJours, colJour).value = 4 Then
                found = True
                Exit For
            End If
        End If
    Next colJour
    
    If Not found Then
        MsgBox "Impossible de trouver le jour '4' dans la ligne " & ligneJours, vbExclamation
        Exit Sub
    End If
    
    ' 3. Lire la couleur de la cellule
    colorDetectee = ws.Cells(rFound.row, colJour).Interior.Color
    
    ' Verifier que ce n'est pas Blanc (16777215) ou Vide
    If colorDetectee = 16777215 Or colorDetectee = 0 Then
        MsgBox "La cellule du 4 pour " & nomPersonne & " est blanche ou sans couleur !", vbExclamation
        Exit Sub
    End If
    
    ' 4. Mettre a jour Feuil_Config
    UpdateConfigColor "COULEUR_BLEU_CLAIR", colorDetectee, "Bleu detecte auto (" & nomPersonne & ")"
    
    
    ' --- PARTIE 2 : DETECTION JAUNE (Aurore) ---
    nomPersonne = "Bourgeois_Aurore"
    Set rFound = ws.Columns("A:B").Find(nomPersonne, LookIn:=xlValues, LookAt:=xlPart)
    
    If Not rFound Is Nothing Then
        ' Chercher le jour 3 (Mardi 3 Mars) - souvent Jaune pour Aurore
        found = False
        For colJour = 3 To 35
             If IsNumeric(ws.Cells(ligneJours, colJour).value) Then
                If ws.Cells(ligneJours, colJour).value = 3 Then
                    found = True
                    Exit For
                End If
            End If
        Next colJour
        
        If found Then
            Dim colorJaune As Long
            colorJaune = ws.Cells(rFound.row, colJour).Interior.Color
             If colorJaune <> 16777215 And colorJaune <> 0 Then
                UpdateConfigColor "COULEUR_INF_ADMIN", colorJaune, "Jaune detecte auto (Aurore)"
                MsgBox "Couleur JAUNE detectee : " & colorJaune, vbInformation
             End If
        End If
    End If

    ' 5. Relancer le calcul
    MsgBox "Configuration des couleurs mise a jour !" & vbCrLf & _
           "Lancement du recalcul...", vbInformation
           
    Run "Calculer_Totaux_Planning"
    
End Sub

Private Sub UpdateConfigColor(key As String, colorVal As Long, comment As String)
    Dim wsCfg As Worksheet
    On Error Resume Next
    Set wsCfg = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    If wsCfg Is Nothing Then Exit Sub
    
    ' Chercher la clé
    Dim rKey As Range
    Set rKey = wsCfg.Columns("A").Find(key, LookIn:=xlValues, LookAt:=xlWhole)
    
    If rKey Is Nothing Then
        ' Ajouter a la fin
        Dim lr As Long
        lr = wsCfg.Cells(wsCfg.Rows.count, "A").End(xlUp).row + 1
        wsCfg.Cells(lr, 1).value = key
        wsCfg.Cells(lr, 2).value = colorVal
        wsCfg.Cells(lr, 3).value = comment
    Else
        rKey.offset(0, 1).value = colorVal
        rKey.offset(0, 2).value = comment
    End If
End Sub
