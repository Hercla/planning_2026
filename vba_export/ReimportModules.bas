Attribute VB_Name = "ReimportModules"
Option Explicit

Public Sub ReimportModules()
    Dim vbProj As Object
    Dim basePath As String
    Dim modules As Variant
    Dim i As Long
    Dim comp As Object
    Dim filePath As String

    ' ?? Nécessite : Trust access to the VBA project object model (Excel Options)
    Set vbProj = ThisWorkbook.VBProject

    basePath = "C:\Users\hercl\planning_2026_repo\vba_export\"

    ' Liste des modules à réimporter (noms EXACTS des composants VBA)
    modules = Array( _
        "Module_Planning_Core", _
        "Module_Remplacement_Auto", _
        "CalculFractionsPresence", _
        "Module_Config_Helpers" _
    )

    On Error GoTo EH

    ' 1) Supprimer les anciens modules (si présents)
    For i = LBound(modules) To UBound(modules)
        Set comp = Nothing
        On Error Resume Next
        Set comp = vbProj.VBComponents(CStr(modules(i)))
        On Error GoTo EH

        If Not comp Is Nothing Then
            vbProj.VBComponents.Remove comp
            Debug.Print "Supprimé: " & CStr(modules(i))
        Else
            Debug.Print "Absent (OK): " & CStr(modules(i))
        End If
    Next i

    ' 2) Importer les nouveaux
    For i = LBound(modules) To UBound(modules)
        filePath = basePath & CStr(modules(i)) & ".bas"

        If Len(Dir$(filePath)) = 0 Then
            Debug.Print "Fichier introuvable: " & filePath
        Else
            vbProj.VBComponents.Import filePath
            Debug.Print "Importé: " & CStr(modules(i))
        End If
    Next i

    MsgBox "Import terminé. Fais : Debug > Compile VBAProject", vbInformation
    Exit Sub

EH:
    MsgBox "Erreur " & Err.Number & " : " & Err.description, vbCritical, "ReimportModules"
End Sub
Sub ReimportFix()
      Dim vbProj As Object, comp As Object
      Set vbProj = ThisWorkbook.VBProject

      On Error Resume Next
      Set comp = vbProj.VBComponents("Module_Remplacement_Auto")
      If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
      vbProj.VBComponents.Import "C:\Users\hercl\planning_2026_repo\vba_export\Module_Remplacement_Auto.bas"
      On Error GoTo 0

      MsgBox "Done. Compiler maintenant.", vbInformation
  End Sub
Sub FixClsModules()
      Dim vbProj As Object, comp As Object
      Dim modules As Variant, i As Long
      Dim basePath As String

      Set vbProj = ThisWorkbook.VBProject
      basePath = "C:\Users\hercl\planning_2026_repo\vba_export\"

      modules = Array("Sept1", "Aout1", "Avril1", "Juillet1", "Mai11", "Nov1", "Oct1", "Decembre2", "Mars1")

      On Error Resume Next
      For i = LBound(modules) To UBound(modules)
          Set comp = vbProj.VBComponents(modules(i))
          If Not comp Is Nothing Then
              vbProj.VBComponents.Remove comp
          End If
          vbProj.VBComponents.Import basePath & modules(i) & ".cls"
          Debug.Print "Fixed: " & modules(i)
      Next i
      On Error GoTo 0

      MsgBox "Done. Compiler.", vbInformation
  End Sub
Sub FixModule()
      Dim vbProj As Object
      Set vbProj = ThisWorkbook.VBProject
      On Error Resume Next
      vbProj.VBComponents.Remove vbProj.VBComponents("Module_Fix_Couleurs")
      vbProj.VBComponents.Import "C:\Users\hercl\planning_2026_repo\vba_export\Module_Fix_Couleurs.bas"
      On Error GoTo 0
      MsgBox "Done"
  End Sub
 Sub ExportAllVBA()
      Dim vbComp As Object
      Dim basePath As String
      basePath = "C:\Users\hercl\planning_2026_repo\vba_export\"

      For Each vbComp In ThisWorkbook.VBProject.VBComponents
          Select Case vbComp.Type
              Case 1 ' Module
                  vbComp.Export basePath & vbComp.Name & ".bas"
              Case 2 ' Class
                  vbComp.Export basePath & vbComp.Name & ".cls"
              Case 3 ' Form
                  vbComp.Export basePath & vbComp.Name & ".frm"
          End Select
      Next
      MsgBox "Export terminé"
  End Sub
