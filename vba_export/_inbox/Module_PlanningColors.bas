' ExportedAt: 2026-01-13 15:00:00 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module_PlanningColors"
Option Explicit

' ===================================================================================
' MODULE :          Module_PlanningColors
' DESCRIPTION :     Fonctions utilitaires pour les couleurs du planning
'                   Lit les couleurs depuis tblCFG (format Long)
' ===================================================================================

' Cache pour les couleurs (evite de relire config a chaque appel)
Private m_ColorWeekend As Long
Private m_ColorFerie As Long
Private m_ColorPoliceWE As Long
Private m_ColorPoliceFerie As Long
Private m_ColorsLoaded As Boolean

' ===================================================================================
' CHARGER LES COULEURS DEPUIS CONFIG
' ===================================================================================

Public Sub LoadPlanningColors()
    On Error Resume Next
    
    m_ColorWeekend = CLng(CfgValueOr("PLN_Couleur_Weekend", 15773696))    ' Bleu clair
    m_ColorFerie = CLng(CfgValueOr("PLN_Couleur_Ferie", 255))             ' Rouge
    m_ColorPoliceWE = CLng(CfgValueOr("PLN_Couleur_Police_Weekend", 16777215))  ' Blanc
    m_ColorPoliceFerie = CLng(CfgValueOr("PLN_Couleur_Police_Ferie", 16777215)) ' Blanc
    
    m_ColorsLoaded = True
    On Error GoTo 0
End Sub

Private Function CfgValueOr(key As String, defaultVal As Long) As Long
    Dim s As String
    s = CfgTextOr(key, "")
    If IsNumeric(s) Then
        CfgValueOr = CLng(s)
    Else
        CfgValueOr = defaultVal
    End If
End Function

' ===================================================================================
' ACCESSEURS PUBLICS
' ===================================================================================

Public Function GetColorWeekend() As Long
    If Not m_ColorsLoaded Then LoadPlanningColors
    GetColorWeekend = m_ColorWeekend
End Function

Public Function GetColorFerie() As Long
    If Not m_ColorsLoaded Then LoadPlanningColors
    GetColorFerie = m_ColorFerie
End Function

Public Function GetColorPoliceWeekend() As Long
    If Not m_ColorsLoaded Then LoadPlanningColors
    GetColorPoliceWeekend = m_ColorPoliceWE
End Function

Public Function GetColorPoliceFerie() As Long
    If Not m_ColorsLoaded Then LoadPlanningColors
    GetColorPoliceFerie = m_ColorPoliceFerie
End Function

' ===================================================================================
' APPLIQUER COULEURS A UNE PLAGE
' ===================================================================================

Public Sub ColorierWeekend(rng As Range)
    If Not m_ColorsLoaded Then LoadPlanningColors
    With rng
        .Interior.Color = m_ColorWeekend
        .Font.Color = m_ColorPoliceWE
        .Font.Bold = True
    End With
End Sub

Public Sub ColorierFerie(rng As Range)
    If Not m_ColorsLoaded Then LoadPlanningColors
    With rng
        .Interior.Color = m_ColorFerie
        .Font.Color = m_ColorPoliceFerie
        .Font.Bold = True
    End With
End Sub

Public Sub ColorierWeekendOuFerie(rng As Range, isHoliday As Boolean)
    If isHoliday Then
        ColorierFerie rng
    Else
        ColorierWeekend rng
    End If
End Sub

' ===================================================================================
' FORCER RELOAD (si config modifiee)
' ===================================================================================

Public Sub ReloadColors()
    m_ColorsLoaded = False
    LoadPlanningColors
End Sub
