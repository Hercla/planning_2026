Attribute VB_Name = "Module_Debug_Couleurs"
Option Explicit

Sub Debug_Couleur_Cellule()
    ' Affiche le code couleur de la cellule sélectionnée
    Dim c As Range
    Set c = Selection
    
    Dim msg As String
    msg = "=== COULEUR CELLULE ===" & vbCrLf
    msg = msg & "Adresse: " & c.Address & vbCrLf
    msg = msg & "Valeur: " & c.value & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "Couleur fond (Interior.Color): " & c.Interior.Color & vbCrLf
    msg = msg & "ColorIndex: " & c.Interior.ColorIndex & vbCrLf
    msg = msg & vbCrLf
    
    ' Convertir en RGB
    Dim colorVal As Long
    colorVal = c.Interior.Color
    msg = msg & "RGB: (" & (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & ((colorVal \ 65536) Mod 256) & ")" & vbCrLf
    
    MsgBox msg, vbInformation, "Debug Couleur"
End Sub
