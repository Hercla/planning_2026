VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Note de Remplacement"
   ClientHeight    =   1980
   ClientLeft      =   -182
   ClientTop       =   -959
   ClientWidth     =   5040
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ===============================================
' === CODE À METTRE DANS LE USERFORM4           ===
' ===============================================
Option Explicit

' Variable pour savoir si l'utilisateur a cliqué sur Annuler

Public WasCancelled As Boolean

' Au chargement du formulaire, on remplit la liste déroulante
Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    Dim tbl As ListObject
    Dim tblRange As Range
    Dim cell As Range

    WasCancelled = True ' Par défaut, on considère que c'est annulé

    On Error GoTo ErrorHandler
    ' On se réfère à la feuille "Remplacants" et au tableau "T_Remplacants"
    Set sh = ThisWorkbook.Sheets("Remplacants")
    Set tbl = sh.ListObjects("T_Remplacants")
    
    Set tblRange = tbl.ListColumns(1).DataBodyRange
    
    Me.cmbNom.Clear
    
    For Each cell In tblRange.Cells
        Me.cmbNom.AddItem cell.value
    Next cell
    
    ' --- AJOUT POUR LE POSITIONNEMENT ---
    ' Positionne le formulaire à droite de la cellule active
    Me.Top = Application.ActiveCell.Top + 20 ' Un peu en dessous du haut de la cellule
    Me.Left = Application.ActiveCell.Left + Application.ActiveCell.width + 5 ' Juste à droite de la cellule
    ' --- FIN DE L'AJOUT ---
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors du chargement de la liste." & vbCrLf & _
           "Vérifiez que la feuille 'Remplacants' et le tableau 'T_Remplacants' existent bien.", _
           vbCritical, "Erreur de Configuration"
    Unload Me
End Sub

' Clic sur le bouton OK
Private Sub cmdOK_Click()
    If Me.cmbNom.value = "" Then
        MsgBox "Veuillez sélectionner un nom dans la liste.", vbExclamation, "Nom requis"
        Exit Sub
    End If
    
    WasCancelled = False
    Me.Hide
End Sub

' Clic sur le bouton Annuler
Private Sub cmdCancel_Click()
    Unload Me
End Sub

' Permet de fermer proprement avec la croix rouge
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub

