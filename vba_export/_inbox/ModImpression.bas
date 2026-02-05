' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "ModImpression"
Option Explicit ' Bonne pratique : force la déclaration de toutes les variables

Sub ImprimerPage1FeuilleActive()

    ' --- Explication ---
    ' Cette macro imprime UNIQUEMENT la page 1 de la feuille actuellement active.
    ' Elle utilise l'imprimante configurée par défaut dans Windows.
    ' Elle N'AFFICHE PAS la boîte de dialogue d'impression.

    Dim nomFeuille As String
    
    ' Vérifier qu'une feuille de calcul est bien active
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Veuillez sélectionner une feuille de calcul avant de lancer l'impression.", vbExclamation, "Action requise"
        Exit Sub
    End If
    
    nomFeuille = ActiveSheet.Name ' Stocker le nom pour le message

    On Error GoTo GestionErreur ' Prévoir le cas où l'impression échoue (imprimante non dispo, etc.)

    ' --- Impression directe de la page 1 ---
    ' From:=1 : Commence à la page 1
    ' To:=1   : Finit à la page 1
    ' Copies:=1 : Imprime une seule copie
    ' Collate:=True : Assembler (pertinent si copies > 1, mais bonne pratique de le laisser)
    ' Preview:=False : Ne pas afficher l'aperçu avant impression
    ' ActivePrinter:= (Optionnel) Vous pourriez spécifier une imprimante ici si besoin.
    '                 Ex: ActivePrinter:="\\Serveur\NomImprimante sur Ne01:"
    '                 Si omis, utilise l'imprimante par défaut de Windows.
    ActiveSheet.PrintOut From:=1, To:=1, Copies:=1, Collate:=True, Preview:=False

    ' Si on arrive ici, l'ordre d'impression a été envoyé (pas de garantie qu'il soit sorti de l'imprimante)
    On Error GoTo 0 ' Désactiver la gestion d'erreur spécifique

    MsgBox "La page 1 de la feuille '" & nomFeuille & "' a été envoyée à l'imprimante par défaut.", vbInformation, "Impression Page 1 Lancée"

    Exit Sub ' Termine la procédure normalement

GestionErreur:
    ' Une erreur s'est produite pendant l'appel à PrintOut
    MsgBox "Impossible de lancer l'impression de la page 1 de la feuille '" & nomFeuille & "'." & vbCrLf & vbCrLf & _
           "Vérifiez que l'imprimante par défaut est bien configurée, connectée et prête." & vbCrLf & _
           "(Erreur VBA: " & Err.Number & " - " & Err.Description & ")", vbCritical, "Erreur d'impression"
    On Error GoTo 0 ' Important pour réinitialiser la gestion d'erreur interne de VBA

End Sub

