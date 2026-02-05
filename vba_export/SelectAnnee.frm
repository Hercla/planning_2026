VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectAnnee 
   Caption         =   "Choix de l'année"
   ClientHeight    =   3120
   ClientLeft      =   -63
   ClientTop       =   -182
   ClientWidth     =   4851
   OleObjectBlob   =   "SelectAnnee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectAnnee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub OK_Click()
'je me suis inspiré d'une macro de Thierry (un autre)
'disponible vers ce lien sur le "Forum Excel Download"
'http://www.excel-downloads.com/html/French/forum/read.php?f=1&i=7120&t=7077
'j'ai simplement modifié certains paramètres pour l'adapter à ce programme
'Pourquoi s'fatiguer alors que tout existe dans ce forum...lol
Dim annee As Integer
Sheets("Parametres").Select
Range("DATE").value = ComboBox1.value
On Error GoTo vide
annee = ComboBox1.value
        If annee < 2003 Then
            MsgBox "L'Annee ne peut être inférieure à 2003"
            ComboBox1 = "2003"
            ComboBox1.SetFocus
        Exit Sub
        End If
        
        If annee = 0 Then
            MsgBox "L'année 0 est dépassée depuis bien longtemps... Le saviez-vous? :o))"
            ComboBox1 = "2003"
            ComboBox1.SetFocus
        Exit Sub
        End If
' ici c'est pour rire...lol
        If annee > 2020 Then
            MsgBox "Désolé, votre ordinateur estime qu'en 2020, vous ne travaillerez plus ici."
            ComboBox1 = ""
            ComboBox1.SetFocus
        Exit Sub
        End If
' se place en A1
' ferme l'userform
Unload SelectAnnee
Sheets("Parametres").Select
Range("E6").Select

'regarde si année bissextile
Bissextile
Exit Sub
vide:
    If ComboBox1 <> "" Then
    MsgBox "Allez directement en prison! Ne passez pas par la case départ...recommencez"
    ComboBox1 = "2003"
End If
If ComboBox1 = "" Then
    MsgBox "Un petit effort, entrez un année!"
    ComboBox1 = "2003"
End If
Exit Sub
End Sub

Private Sub Annuler_Click()
Unload Me
Sheets("Config_Calendrier").Select
Range("C2").Select
End Sub
' ici on propose une liste d'année
' dans la comboBox
Private Sub UserForm_Initialize()
ComboBox1.value = Range("DATE").value
Dim i As Integer
For i = 2003 To 2020
ComboBox1.AddItem i
Next i
End Sub

