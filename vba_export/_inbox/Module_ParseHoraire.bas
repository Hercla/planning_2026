Attribute VB_Name = "Module_ParseHoraire"
Option Explicit

' =========================================================================================
'   MODULE PARSE HORAIRE - Fonctions Utilitaires
'   Gère le parsing des codes, conversions d'heures et calculs de présences
' =========================================================================================

Public Function HeureEnDecimal(heureStr As String) As Double
    ' Convertit "HH:MM" ou "HH" en décimal (ex: "6:45" -> 6.75)
    Dim parts() As String
    If InStr(heureStr, ":") > 0 Then
        parts = Split(heureStr, ":")
        HeureEnDecimal = CDbl(parts(0)) + (CDbl(parts(1)) / 60)
    ElseIf IsNumeric(heureStr) Then
        HeureEnDecimal = CDbl(heureStr)
    Else
        HeureEnDecimal = 0
    End If
End Function

Public Function ParseCodeHoraire(code As String, _
                               ByRef start1 As Double, ByRef end1 As Double, _
                               ByRef start2 As Double, ByRef end2 As Double) As Boolean
    ' Parse un code horaire simple ou coupé
    ' Ex: "6:45 15:15" -> start1=6.75, end1=15.25
    ' Ex: "8 12 14 18" -> start1=8, end1=12, start2=14, end2=18
    
    start1 = 0: end1 = 0: start2 = 0: end2 = 0
    ParseCodeHoraire = False
    
    Dim parts() As String
    ' Remplacer les espaces multiples et sauts de ligne
    code = Replace(Trim(code), vbLf, " ")
    code = Replace(code, vbCr, " ")
    Do While InStr(code, "  ") > 0
        code = Replace(code, "  ", " ")
    Loop
    
    parts = Split(code, " ")
    
    On Error GoTo ErrParse
    
    If UBound(parts) = 1 Then
        ' Code simple: Début Fin
        start1 = HeureEnDecimal(parts(0))
        end1 = HeureEnDecimal(parts(1))
        ParseCodeHoraire = True
        
    ElseIf UBound(parts) >= 3 Then
        ' Code coupé: Début1 Fin1 Début2 Fin2
        start1 = HeureEnDecimal(parts(0))
        end1 = HeureEnDecimal(parts(1))
        start2 = HeureEnDecimal(parts(2))
        end2 = HeureEnDecimal(parts(3))
        ParseCodeHoraire = True
    End If
    Exit Function

ErrParse:
    ParseCodeHoraire = False
End Function

Public Function CalculerPresenceCreneau(hDebut As Double, hFin As Double, _
                                      targetDebut As Double, targetFin As Double) As Double
    ' Calcule la fraction de présence dans un créneau cible
    Dim overlapStart As Double, overlapEnd As Double
    
    overlapStart = Application.Max(hDebut, targetDebut)
    overlapEnd = Application.Min(hFin, targetFin)
    
    If overlapEnd > overlapStart Then
        CalculerPresenceCreneau = overlapEnd - overlapStart
    Else
        CalculerPresenceCreneau = 0
    End If
End Function

Public Function EstPresentA(hDebut As Double, hFin As Double, hCible As Double) As Boolean
    ' Vérifie si une personne est présente à une heure précise (avec tolérance)
    Const TOLERANCE As Double = 0.001
    EstPresentA = (hDebut <= hCible + TOLERANCE And hFin > hCible + TOLERANCE)
End Function

Public Sub CalculerPresencesPeriodes(h1 As Double, f1 As Double, h2 As Double, f2 As Double, _
                                     ByRef matin As Double, ByRef am As Double, _
                                     ByRef soir As Double, ByRef nuit As Double)
    ' Règles Périodes (en heures décimales)
    ' Matin: 6h-13h30 (6.0 - 13.5)
    ' AM: 13h30-21h (13.5 - 21.0)
    ' Soir: fraction après 19h (19.0) ??? -> Règle standard: >19h ou spécifique ?
    ' Nuit: 21h-6h (21.0 - 6.0 du lendemain)
    
    ' Note : Les règles exactes dépendent de votre configuration "CalculFractionsPresence"
    ' Ici j'applique une logique standard basée sur vos fichiers précédents
    
    Dim total As Double
    total = (f1 - h1) + (f2 - h2)
    
    ' Logique simplifiée basée sur les codes existants:
    ' Matin : Présence significative avant 13h30
    If CalculerPresenceCreneau(h1, f1, 6, 13.5) + CalculerPresenceCreneau(h2, f2, 6, 13.5) > 1 Then
        matin = 1
    End If
    
    ' AM : Présence significative après 13h30
    If CalculerPresenceCreneau(h1, f1, 13.5, 21) + CalculerPresenceCreneau(h2, f2, 13.5, 21) > 1 Then
        am = 1
    End If
    
    ' Soir : Souvent défini comme présent après 19h ou 20h
    If f1 > 19 Or f2 > 19 Then
        soir = 1
    End If
    
    ' Nuit : Présence après 21h ou avant 6h
    If f1 > 21 Or f2 > 21 Or h1 < 6 Then ' Simplifié
        nuit = 1
    End If
End Sub

Public Sub CalculerPresencesSpecifiques(hd As Double, hf As Double, hd2 As Double, hf2 As Double, _
                                        ByRef p0645 As Long, ByRef p7h8h As Long, ByRef p8h1630 As Long)
    ' P_0645: Commence à 6h45 pile (tolérance stricte)
    If Abs(hd - 6.75) < 0.01 Then p0645 = 1 Else p0645 = 0
    
    ' P_7H8H: Présent complet entre 7h et 8h
    If EstPresentA(hd, hf, 7) And EstPresentA(hd, hf, 7.9) Then
        p7h8h = 3 ' Valeur 3 selon règle
    Else
        p7h8h = 0
    End If
    
    ' P_8H1630: Couvre 8h à 16h30
    If hd <= 8 And hf >= 16.5 Then
        p8h1630 = 1
    Else
        p8h1630 = 0
    End If
End Sub

Public Function EstHoraireSpecial(h1 As Double, f1 As Double, h2 As Double, f2 As Double, defStr As String) As Boolean
    ' Compare les horaires actuels avec une chaîne de définition "H1:M1 H2:M2 H3:M3 H4:M4"
    If defStr = "" Then Exit Function
    
    Dim parts() As String
    parts = Split(defStr, " ")
    
    If UBound(parts) < 3 Then Exit Function ' Besoin de 4 parties pour un coupé
    
    Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double
    c1 = HeureEnDecimal(parts(0))
    c2 = HeureEnDecimal(parts(1))
    c3 = HeureEnDecimal(parts(2))
    c4 = HeureEnDecimal(parts(3))
    
    ' Comparaison avec tolérance
    Const TOL As Double = 0.01
    If Abs(h1 - c1) < TOL And Abs(f1 - c2) < TOL And _
       Abs(h2 - c3) < TOL And Abs(f2 - c4) < TOL Then
        EstHoraireSpecial = True
    End If
End Function
