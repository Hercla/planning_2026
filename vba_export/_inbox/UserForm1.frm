' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "PLANNING TEAM US 1D"
   ClientHeight    =   10185
   ClientLeft      =   -1155
   ClientTop       =   -5880
   ClientWidth     =   3703
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox3_Change()
    SetPlanningCellFromConfig "W13"
End Sub

Private Sub ComboBox4_Change()
End Sub
Private Sub CommandButton24_Click()
UserForm5.Show
End Sub

Private Sub CommandButton25_Click()
UserForm5.Show
End Sub

Private Sub CommandButton27_Click()
Sheets("AIDE").Activate
End Sub


Private Sub ComboBox1_Change()
End Sub

Private Sub CommandButton1_Click()
Sheets("PLANNING").Activate
Selection.AutoFilter Field:=2, Criteria1:="REEL"
End Sub

Private Sub CommandButton10_Click()
Sheets("Mai").Activate
Unload UserForm1
End Sub

Private Sub CommandButton11_Click()
Sheets("Juin").Activate
Unload UserForm1
End Sub

Private Sub CommandButton12_Click()
Sheets("Juillet").Activate
Unload UserForm1
End Sub

Private Sub CommandButton13_Click()
Sheets("Aout").Activate
Unload UserForm1
End Sub

Private Sub CommandButton14_Click()
Sheets("Sept").Activate
Unload UserForm1
End Sub

Private Sub CommandButton15_Click()
Sheets("Oct").Activate
Unload UserForm1
End Sub

Private Sub CommandButton16_Click()
Sheets("Nov").Activate
Unload UserForm1
End Sub

Private Sub CommandButton17_Click()
Sheets("Dec").Activate
Unload UserForm1
End Sub

Private Sub CommandButton2_Click()
Sheets("PLANNING").Activate
Selection.AutoFilter Field:=2, Criteria1:="PREV"
    Range("B29").Select
End Sub

Private Sub ComboBox5_Change()
ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton100_Click()
    SetPlanningCellFromConfig "W29"
End Sub

Private Sub CommandButton101_Click()
    SetPlanningCellFromConfig "W34"
End Sub
Private Sub btnMajDates_Click()
On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1) Génère jours + numéros + couleurs sur Janv..Dec
    Call GenerateurCalendrier.GenererDatesEtJoursPourTousLesMois

    ' 2) Met à jour Config_Codes (F/R + heures) selon Feuil_Config!B2
    Call MettreAJourConfigurationCodes

    MsgBox "MAJ DATES terminée : calendriers + fériés (F/R) + heures codes.", vbInformation
    GoTo CleanExit

ErrHandler:
    MsgBox "Erreur MAJ DATES : " & Err.Number & " - " & Err.Description, vbCritical

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub CommandButton102_Click()
'coller code CTR et l adapter par rapport au mois -1
Module_UserActions.InsertCodeFromUserForm "CTR"
End Sub

Private Sub CommandButton103_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "14 20"
End Sub

Private Sub CommandButton104_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "13 19"
End Sub

Private Sub CommandButton105_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "12:30 16:30"
End Sub

Private Sub CommandButton107_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "10 19"
End Sub

Private Sub CommandButton108_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "9 15:30"
End Sub

Private Sub CommandButton109_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "8 16:30"
End Sub

Private Sub CommandButton110_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "8 14"
End Sub

Private Sub CommandButton111_Click()
'place un astérix si besoin inf niveau du code
Call ToggleAsterisqueCellule
End Sub

Private Sub CommandButton112_Click()
'met la couleur vert foncé code hrel
Call ToggleColorierCelluleVertFonce
End Sub

Private Sub CommandButton113_Click()
    SetPlanningCellFromConfig "W34", "Arial Narrow", 8
End Sub

Private Sub CommandButton114_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 15 di"
End Sub

Private Sub CommandButton115_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 15"
End Sub

Private Sub CommandButton116_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 20"
End Sub

Private Sub CommandButton117_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 19 di"
End Sub

Private Sub CommandButton118_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 19"
End Sub

Private Sub CommandButton119_Click()
' masquer toutes les lignes de "43:50" en premier lieu. Ensuite, la boucle parcourt chaque ligne et la
'rend visible seulement si elle est vide et que le nombre de lignes visibles est inférieur au nombre demandé.
Call AfficherMasquerLignesDynamiques

End Sub

Private Sub CommandButton120_Click()
    ' Afficher le UserForm pour la mise à jour des lignes
    Call MajLigne
End Sub


Private Sub CommandButton121_Click()
    SetPlanningCellValue "TV", RGB(255, 255, 0), RGB(0, 0, 0)
End Sub

Private Sub CommandButton125_Click()
'place une note
Call CreerNoteRemplacement
End Sub

Private Sub CommandButton126_Click()
Call GenererRoulementOptimise
End Sub

Private Sub CommandButton127_Click()
Call ImprimerPage1FeuilleActive

End Sub

Private Sub CommandButton128_Click()
Call CTR_CheckWeekendEligibility

End Sub

Private Sub CommandButton129_Click()
Call CheckDPMonthlyCodes

End Sub

Private Sub CommandButton130_Click()
Call UpdateMonthlySheets_Final_Polished

End Sub

Private Sub CommandButton131_Click()
Call Check_Presence_Infirmiers
End Sub

Private Sub CommandButton132_Click()
Call AnalyseEtRemplacementPlanningUltraOptimise

End Sub

Private Sub CommandButtonNui_Click()
    UX_Mode_Nuit_ActiveSheet
End Sub
Private Sub CommandButton28_Click()
    Sheets("HORAIRES").Activate
    Range("A1:J1").Select
    ActiveWindow.Zoom = 100
    Range("C5").Select
End Sub

Private Sub CommandButton29_Click()
    Sheets("CYCLES").Activate
    Range("A1:AT1").Select
    ActiveWindow.Zoom = 100
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

Private Sub CommandButton5_Click()
Sheets("AIDE").Activate
Unload UserForm1
End Sub

Private Sub CommandButton6_Click()
Sheets("Janv").Activate
Unload UserForm1
End Sub

Private Sub CommandButton7_Click()
Sheets("Fev").Activate
Unload UserForm1
End Sub

Private Sub CommandButton8_Click()
Sheets("Mars").Activate
Unload UserForm1
End Sub

Private Sub CommandButton9_Click()
Sheets("04").Activate
Unload UserForm1
End Sub


Private Sub CommandButton23_Click()
Unload UserForm1
End Sub
Private Sub CommandButton33_Click()
Unload UserForm1
End Sub

Private Sub CommandButton31_Click()
Call RAZPlanMens
End Sub

Private Sub CommandButton32_Click()
Call GenererRoulement8SemJusqu31Dec_DynamiqueNuit

End Sub

Private Sub CommandButton34_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Config_Calendrier"
End Sub
Private Sub CommandButton36_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Janv"
End Sub

Private Sub CommandButton37_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Fev"
End Sub

Private Sub CommandButton38_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Mars"
End Sub

Private Sub CommandButton39_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Avril"
End Sub

Private Sub CommandButton40_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Mai"
End Sub

Private Sub CommandButton41_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Juin"
End Sub

Private Sub CommandButton42_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Dec"
End Sub

Private Sub CommandButton43_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Nov"
End Sub

Private Sub CommandButton44_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Oct"
End Sub

Private Sub CommandButton45_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Sept"
End Sub

Private Sub CommandButton46_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Aout"
End Sub

Private Sub CommandButton47_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Juillet"
End Sub
Private Sub CommandButton48_Click() ' Le bouton pour "6:45 12:45"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "6:45 12:45"
End Sub

Private Sub CommandButton49_Click() ' Le bouton pour "6:45 15:15"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "6:45 15:15"
End Sub
Private Sub CommandButton50_Click() ' Le bouton pour "7 15:30"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7 15:30"
End Sub

Private Sub CommandButton52_Click() ' Le bouton pour "8 16:30"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "8 16:30"
End Sub

Private Sub CommandButton53_Click() ' Le bouton pour "7 13"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7 13"
End Sub

Private Sub CommandButton55_Click() ' Le bouton pour "9 15:30"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "9 15:30"
End Sub

Private Sub CommandButton56_Click() ' Le bouton pour "10 20"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "10 20"
End Sub

Private Sub CommandButton57_Click() ' Le bouton pour "13 19"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "13 19"
End Sub

Private Sub CommandButton58_Click() ' Le bouton pour "14 20"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "14 20"
End Sub

Private Sub CommandButton60_Click() ' Le bouton pour "7:15 13:15"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7:15 13:15"
End Sub
Private Sub CommandButton61_Click() ' Le bouton pour "C 19"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 19"
End Sub

Private Sub CommandButton62_Click() ' Le bouton pour "C 19 di"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 19 di"
End Sub

Private Sub CommandButton63_Click() ' Le bouton pour "C 20"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 20"
End Sub

Private Sub CommandButton64_Click() ' Le bouton pour "C 15"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 15"
End Sub

Private Sub CommandButton65_Click() ' Le bouton pour "C 15 di"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 15 di"
End Sub

Private Sub CommandButton66_Click() ' Le bouton pour "19:45 06:45"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "19:45 06:45"
End Sub

Private Sub CommandButton67_Click() ' Le bouton pour "20 7"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "20 7"
End Sub

Private Sub CommandButton69_Click() ' Le bouton pour "20 7"
'Sub bouton_Rhs 8h
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "20 7"
End Sub

Private Sub CommandButton70_Click() ' Le bouton pour "20 7"
'Sub bouton_CSOC
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "20 7"
End Sub

Private Sub CommandButton71_Click() ' Le bouton pour "20 7"
'Sub bouton_CSOC
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "20 7"
End Sub

Private Sub CommandButton72_Click() ' Le bouton pour "7:15 15:45"
'Sub bouton_21()7:15 15:45
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7:15 15:45"
End Sub

Private Sub CommandButton73_Click() ' Le bouton pour "4/5*"
'Sub bouton_23()"4/5*"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "4/5*"
End Sub

Private Sub CommandButton74_Click() ' Le bouton pour "3/4*"
'Sub bouton_22()"3/4*"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "3/4*"
End Sub

Private Sub CommandButton75_Click()
    SetPlanningCellFromConfig "W32", "Arial", 12, RGB(0, 112, 192), RGB(255, 255, 255)
End Sub

Private Sub CommandButton82_Click()
'Sub bouton_22() 7 11
InsertCodeAndMove "7 11"
End Sub

Private Sub CommandButton83_Click()
'Sub bouton_22() 8 14
InsertCodeAndMove "8 14"
End Sub

Private Sub CommandButton84_Click()
'Sub bouton_22() RHS 6h
InsertCodeAndMove "RHS 6h"
End Sub

Private Sub CommandButton85_Click()
' 'Sub bouton_22() RHS 6h
InsertCodeAndMove "RHS 6h"
End Sub

Private Sub CommandButton86_Click()
' 'Sub bouton_22() RHS 6h
InsertCodeAndMove "RHS 6h"
End Sub

Private Sub CommandButton87_Click()
Call ModulePDFGeneration.Generate_PDF_Jour
End Sub

Private Sub CommandButton88_Click()
Call ModulePDFGeneration.Generate_PDF_Nuit
End Sub

Private Sub CommandButton89_Click()
    Call Mode_Jour
End Sub
Sub AfficherMasquerLignes()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lignes As Range
    Set lignes = ws.Rows("43:44")
    
    If lignes.Hidden = True Then
        lignes.Hidden = False ' Affiche les lignes 43 et 44
    Else
        lignes.Hidden = True ' Masque les lignes 43 et 44
    End If
End Sub

Private Sub AfficherMasquerLignesDynamiques()
    ' TODO: remplacer par la vraie logique si besoin
    AfficherMasquerLignes
End Sub
Private Sub CommandButton90_Click()
Call ColorCells2
End Sub

Public Sub CommandButton91_Click()
Call Mode_Nuit
ActiveWindow.Zoom = 70 'Réglage de zoom
End Sub

    Private Sub CommandButtonJour_Click()
    UX_Mode_Jour_ActiveSheet
End Sub

Private Sub CommandButton93_Click()
    ' Appelle la macro publique qui se trouve dans Module_Planning
    Call UpdateDailyTotals_V2
End Sub


Private Sub CommandButton94_Click()
' Call the PasteToPlanning macro
    PasteToPlanning
End Sub

Private Sub CommandButton96_Click()
' Call colorier cellule verte macro pr hrel
Call ToggleColorierCelluleVertFonce
End Sub

Private Sub CommandButton95_Click()
    FormulaireEntrees.Show
End Sub

Private Sub CommandButton97_Click()
' Call colorier cellule verte macro pr 7 15 30 asbd
Call ToggleColorierCelluleBleuClair
End Sub

Private Sub CommandButton98_Click()
'place un astérix si besoin inf niveau du code
Call ToggleAsterisqueCellule

End Sub

Private Sub CommandButton99_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7:30 16"
End Sub

Private Sub Frame2_Click()

End Sub



Private Function IsPlanningCell() As Boolean
    On Error Resume Next
    IsPlanningCell = Not Intersect(ActiveCell, ActiveSheet.Range("planning")) Is Nothing
    On Error GoTo 0
End Function

Private Sub SetPlanningCellFromConfig(ByVal cfgAddr As String, _
                                      Optional ByVal fontName As String = "", _
                                      Optional ByVal fontSize As Variant = Empty, _
                                      Optional ByVal bgColor As Variant = Empty, _
                                      Optional ByVal fontColor As Variant = Empty)
    If Not IsPlanningCell() Then Exit Sub

    With ActiveCell
        .Value = Sheets("Config_Calendrier").Range(cfgAddr).Value
        If IsEmpty(bgColor) Then
            .Interior.Color = RGB(255, 255, 255)
        Else
            .Interior.Color = CLng(bgColor)
        End If
        If IsEmpty(fontColor) Then
            .Font.Color = RGB(0, 0, 0)
        Else
            .Font.Color = CLng(fontColor)
        End If
        If Len(fontName) > 0 Then .Font.Name = fontName
        If Not IsEmpty(fontSize) Then .Font.Size = fontSize
        .Offset(0, 1).Select
    End With
End Sub

Private Sub SetPlanningCellValue(ByVal value As String, ByVal bgColor As Long, ByVal fontColor As Long)
    If Not IsPlanningCell() Then Exit Sub

    With ActiveCell
        .Value = value
        .Interior.Color = bgColor
        .Font.Color = fontColor
        .Offset(0, 1).Select
    End With
End Sub

Private Sub UserForm_Click()
Me.width = 256.55 ' Ajustez la largeur selon vos besoins
End Sub
' --- AJOUTEZ CETTE PROCÉDURE DANS LE CODE DU USERFORM "Menu" ---
Private Sub InsertCodeAndMove(ByVal code As String)
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    
    With ActiveCell
        .value = code
        .Interior.Color = vbWhite
        .Font.Color = vbBlack
        On Error Resume Next
        .offset(0, 1).Select
        On Error GoTo 0
    End With
End Sub
