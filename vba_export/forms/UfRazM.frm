VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UfRazM 
   Caption         =   "Effacement du planning mensuel "
   ClientHeight    =   6960
   ClientLeft      =   -63
   ClientTop       =   -182
   ClientWidth     =   6993
   OleObjectBlob   =   "UfRazM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UfRazM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Sub CBtAnnule_Click()
UfRazM.Hide
End Sub

Sub CBtRazM_Click()
If UfRazM.OBtJanvier = True Then
    Range("D22:G52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("E19").Select
ElseIf UfRazM.OBtFevrier = True Then
    Range("K22:N52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("L19").Select
ElseIf UfRazM.OBtMars = True Then
    Range("R22:U52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("S19").Select
ElseIf UfRazM.OBtAvril = True Then
Range("Y22:AB52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("Z19").Select
ElseIf UfRazM.OBtMai = True Then
Range("AF22:AI52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("AG19").Select
ElseIf UfRazM.OBtJuin = True Then
Range("AM22:AP52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("AN19").Select
ElseIf UfRazM.OBtJuillet = True Then
Range("AT22:AW52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("AU19").Select
ElseIf UfRazM.OBtAout = True Then
Range("BA22:BD52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("BB19").Select
ElseIf UfRazM.OBtSeptembre = True Then
Range("BH22:BK52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("BI19").Select
ElseIf UfRazM.OBtOctobre = True Then
Range("BO22:BR52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("BP19").Select
ElseIf UfRazM.OBtNovembre = True Then
Range("BV22:BY52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("BW19").Select
ElseIf UfRazM.OBtDecembre = True Then
Range("CC22:CF52").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("CD19").Select
ElseIf UfRazM.OBtTous = True Then
Range("CC22:CF52,BV22:BY52,BO22:BR52,BH22:BK52,BA22:BD52,AT22:AW52,AM22:Ap52,AF22:Ai52,Y22:Ab52,R22:U52,K22:N52,D22:G52").Select
Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Range("E19").Select
End If
UfRazM.Hide
End Sub

Sub UserForm_Activate()
With UfRazM
    .LbTxNom = ActiveSheet.Range("D2") & " " & ActiveSheet.Range("D1")
    .OBtJanvier = False
    .OBtFevrier = False
    .OBtMars = False
    .OBtAvril = False
    .OBtMai = False
    .OBtJuin = False
    .OBtJuillet = False
    .OBtAout = False
    .OBtSeptembre = False
    .OBtOctobre = False
    .OBtNovembre = False
    .OBtDecembre = False
    .OBtTous = False
End With
End Sub

