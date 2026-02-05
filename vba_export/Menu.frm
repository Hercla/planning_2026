VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   ClientHeight    =   4095
   ClientLeft      =   -420
   ClientTop       =   -1995
   ClientWidth     =   5733
   OleObjectBlob   =   "Menu.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton10_Click()
    Sheets("Avril").Activate
    Unload Menu
End Sub

Private Sub CommandButton11_Click()
    Sheets("06").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton12_Click()
    Sheets("07").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton13_Click()
    Sheets("08").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton14_Click()
    Sheets("09").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton15_Click()
    Sheets("10").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton16_Click()
    Sheets("11").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton17_Click()
    Sheets("12").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton2_Click()
    Sheets("PLANNING").Activate
    Selection.AutoFilter Field:=2, Criteria1:="PREV"
    Range("B29").Select
End Sub

Private Sub CommandButton21_Click()
End Sub

Private Sub ComboBox5_Change()
    ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton22_Click()
End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub CommandButton28_Click()
    Sheets("HORAIRES").Activate
    Range("A1:J1").Select
    ActiveWindow.Zoom = True
    Range("C5").Select
End Sub

Private Sub CommandButton29_Click()
    Sheets("CYCLES").Activate
    Range("A1:AT1").Select
    ActiveWindow.Zoom = True
    Range("C2").Select
End Sub

Private Sub CommandButton3_Click()
    Sheets("PLANNING").Activate
    Selection.AutoFilter Field:=2
    ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton4_Click()
    Sheets("Config_Calendrier").Activate
    Unload UserForm1
End Sub
Private Sub CommandButton6_Click()
    Sheets("01").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton7_Click()
    Sheets("02").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton8_Click()
    Sheets("03").Activate
    Unload UserForm1
End Sub

Private Sub CommandButton9_Click()
    Sheets("04").Activate
    Unload UserForm1
End Sub
Private Sub CommandButton23_Click()
    Unload UserForm1
End Sub
Private Sub CommandButton26_Click()
    Unload UserForm4
    Sheets("01").Activate
End Sub

Private Sub CommandButton33_Click()
    Unload UserForm1
End Sub
Private Sub CommandButton34_Click()
    Sheets("PARAMETRAGE").Activate
    Range("H1:BB1").Select
    ActiveWindow.Zoom = True
    Range("I6").Select
End Sub

Private Sub CommandButton36_Click()
    'Janv
    Sheets("Janv").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton37_Click()
    'Fev
    Sheets("Fev").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton38_Click()
    'Mars
    Sheets("Mars").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton39_Click()
    'Avril
    Sheets("Avril").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton40_Click()
    'Mai
    Sheets("Mai").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
    
End Sub

Private Sub CommandButton41_Click()
    Sheets("Juin").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
    
End Sub
Private Sub CommandButton47_Click()
    'Juillet
    Sheets("Juillet").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton48_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Config_Calendrier").Range("W2").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.offset(0, 1).Select
End Sub

Private Sub CommandButton77_Click()
    'Aout
    Sheets("Aout").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton45_Click()
    'Septembre
    Sheets("Sept").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
    
End Sub
Private Sub CommandButton76_Click()
    'Octobre
    Sheets("Oct").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton78_Click()
'Novembre
    Sheets("Nov").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton79_Click()
'Décembre
    Sheets("Dec").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub

