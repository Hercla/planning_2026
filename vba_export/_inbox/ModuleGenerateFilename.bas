' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "ModuleGenerateFilename"
Option Explicit

' --- GenerateNewFilename Macro ---
Sub GenerateNewFilename()
    On Error GoTo ErrorHandler

    ' Déclaration des variables
    Dim nomPrenom As String
    Dim currentDate As String
    Dim currentTime As String
    Dim newFilename As String
    Dim postCM As String
    Dim uf As UserFormInputs

    ' Créer une instance du UserForm
    Set uf = New UserFormInputs

    ' Afficher le UserForm de manière modale pour collecter les entrées utilisateur
    uf.Show vbModal

    ' Vérifier si l'utilisateur a annulé
    If uf.IsCancelled Then
        MsgBox "Opération annulée par l'utilisateur.", vbInformation
        Unload uf
        Exit Sub
    End If

    ' Récupérer les données du UserForm
    With uf
        nomPrenom = Trim(.cboNomPrenom.value)

        If .chkPostCM.value = True Then
            postCM = "Post_CM"
        Else
            postCM = ""
        End If
    End With

    ' Récupérer la date et l'heure actuelles avec un format standard
    currentDate = Format(Now, "dd_mm_yyyy")
    currentTime = Format(Now, "hh_mm AM/PM") ' Correction du format

    ' Construire le nouveau nom de fichier
    newFilename = "Demandes_Remplacements_" & postCM & "_De_" & nomPrenom & "_Us_1D_" & currentDate & "_" & currentTime & ".xlsm"

    ' Afficher le nom de fichier généré
    MsgBox "Nom de fichier généré : " & newFilename, vbInformation

    ' Décharger le UserForm
    Unload uf

    Exit Sub

ErrorHandler:
    MsgBox "Une erreur est survenue dans GenerateNewFilename : " & Err.Description, vbCritical
    Resume CleanUp

CleanUp:
    ' Assurez-vous que le UserForm est déchargé en cas d'erreur
    If Not uf Is Nothing Then
        Unload uf
    End If
End Sub

