Attribute VB_Name = "Module_SetYearA1"
Option Explicit

' ============================================================================
' Met l'annee de Feuil_Config dans la cellule A1 de tous les onglets mois
' ============================================================================

Sub SetYearInA1_AllMonths()
    Dim ws As Worksheet
    Dim annee As Variant
    Dim monthSheets As Variant
    Dim shName As Variant

    ' Liste des onglets mois
    monthSheets = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                        "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")

    ' Lire l'annee depuis Feuil_Config
    annee = GetConfigValue("CFG_Year")

    If IsEmpty(annee) Or annee = "" Then
        annee = Year(Date) ' Valeur par defaut si non trouve
    End If

    ' Appliquer a tous les onglets mois
    For Each shName In monthSheets
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(shName))
        If Not ws Is Nothing Then
            ws.Range("A1").Value = annee
            ws.Range("A1").Font.Bold = True
            ws.Range("A1").Font.Size = 14
        End If
        Set ws = Nothing
        On Error GoTo 0
    Next shName

    MsgBox "Annee " & annee & " appliquee dans A1 de tous les onglets mois.", vbInformation
End Sub

Private Function GetConfigValue(key As String) As Variant
    Dim wsConfig As Worksheet
    Dim rng As Range

    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    If wsConfig Is Nothing Then
        GetConfigValue = Empty
        Exit Function
    End If

    Set rng = wsConfig.Columns("A").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        GetConfigValue = rng.Offset(0, 1).Value
    Else
        GetConfigValue = Empty
    End If
    On Error GoTo 0
End Function
