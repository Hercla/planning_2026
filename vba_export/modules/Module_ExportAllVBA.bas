Attribute VB_Name = "Module_ExportAllVBA"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

' ==========================
' CONFIG
' ==========================
Private Const EXPORT_FOLDER_NAME As String = "_EXPORT_VBA"

' ==========================
' ENTRY POINT
' ==========================
Public Sub Export_All_VBA_Modules_Stable()

    Dim baseFolder As String
    Dim targetFolder As String

    baseFolder = GetExportBaseFolder()
    targetFolder = baseFolder & Application.PathSeparator & EXPORT_FOLDER_NAME

    MsgBox _
        "CLASSEUR EXECUTANT LA MACRO :" & vbCrLf & ThisWorkbook.fullName & vbCrLf & vbCrLf & _
        "ThisWorkbook.Path :" & vbCrLf & ThisWorkbook.path & vbCrLf & vbCrLf & _
        "DOSSIER EXPORT FINAL :" & vbCrLf & targetFolder, _
        vbInformation, "TRACE EXPORT VBA"

    EnsureFolderRecursive targetFolder
    CleanExportFolder targetFolder

    ' >>> EXPORT AVEC WRAPPER <<<
    ExportProjectComponents_Safe targetFolder

    CreateRunProofFile targetFolder
    StampExportFiles targetFolder

    MsgBox "? Export VBA TERMINÉ" & vbCrLf & targetFolder, vbInformation

End Sub

' ==========================
' EXPORT CORE (SAFE)
' ==========================
Private Sub ExportProjectComponents_Safe(ByVal targetFolder As String)

    Dim vbProj As Object
    Dim vbComp As Object
    Dim fso As Object
    Dim filePath As String
    Dim ext As String

    ' --- WRAPPER CRITIQUE ---
    On Error GoTo EH_ACCESS
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    ' ------------------------

    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each vbComp In vbProj.VBComponents
        ext = ComponentExtension(vbComp)
        filePath = targetFolder & Application.PathSeparator & SafeFileName(vbComp.Name) & ext

        On Error Resume Next
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
        vbComp.Export filePath
        On Error GoTo 0
    Next vbComp

    Exit Sub

EH_ACCESS:
    MsgBox _
        "? ACCÈS AU PROJET VBA REFUSÉ" & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.description & vbCrLf & vbCrLf & _
        "? Solution :" & vbCrLf & _
        "Excel > Fichier > Options > Centre de gestion de la confidentialité >" & vbCrLf & _
        "Paramètres des macros > cocher :" & vbCrLf & _
        "? Accès approuvé au modèle d’objet du projet VBA" & vbCrLf & vbCrLf & _
        "Puis fermer et rouvrir Excel.", _
        vbCritical, "EXPORT VBA BLOQUÉ"

End Sub

Private Function ComponentExtension(ByVal vbComp As Object) As String
    Select Case vbComp.Type
        Case 1: ComponentExtension = ".bas"
        Case 2: ComponentExtension = ".cls"
        Case 3: ComponentExtension = ".frm"
        Case 100: ComponentExtension = ".cls"
        Case Else: ComponentExtension = ".txt"
    End Select
End Function

' ==========================
' TRACE / PROOF
' ==========================
Private Sub CreateRunProofFile(ByVal targetFolder As String)
    Dim f As Integer, p As String

    p = targetFolder & Application.PathSeparator & _
        "_RUN_PROOF_" & Format(Now, "yyyymmdd_HHNNSS") & ".txt"

    f = FreeFile
    Open p For Output As #f
    Print #f, "Export run proof"
    Print #f, "Timestamp : " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    Print #f, "Workbook  : " & ThisWorkbook.fullName
    Close #f
End Sub

' ==========================
' STAMP FILES
' ==========================
Private Sub StampExportFiles(ByVal folderPath As String)
    Dim fso As Object, f As Object
    Dim ext As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each f In fso.GetFolder(folderPath).Files
        ext = LCase$(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "txt" Then
            PrependStampToTextFile f.path
        End If
    Next f
End Sub

Private Sub PrependStampToTextFile(ByVal filePath As String)
    Dim f As Integer, content As String, stamp As String

    stamp = "' ExportedAt: " & Format(Now, "yyyy-mm-dd HH:nn:ss") & _
            " | Workbook: " & ThisWorkbook.Name & vbCrLf

    f = FreeFile
    Open filePath For Binary Access Read As #f
    content = Space$(LOF(f))
    Get #f, , content
    Close #f

    If Left$(content, 13) = "' ExportedAt:" Then
        content = RemoveFirstLine(content)
    End If

    f = FreeFile
    Open filePath For Binary Access Write As #f
    Put #f, , stamp & content
    Close #f
End Sub

Private Function RemoveFirstLine(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, vbCrLf)
    If p > 0 Then
        RemoveFirstLine = Mid$(s, p + Len(vbCrLf))
    Else
        RemoveFirstLine = s
    End If
End Function

' ==========================
' FILE SYSTEM
' ==========================
Private Sub CleanExportFolder(ByVal folderPath As String)
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each f In fso.GetFolder(folderPath).Files
        Select Case LCase$(fso.GetExtensionName(f.Name))
            Case "bas", "cls", "frm", "txt", "frx", "md"
                On Error Resume Next
                f.Delete True
                On Error GoTo 0
        End Select
    Next f
End Sub

Private Sub EnsureFolderRecursive(ByVal folderPath As String)
    Dim fso As Object, parentPath As String

    If Dir$(folderPath, vbDirectory) <> vbNullString Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    parentPath = fso.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 Then
        If Dir$(parentPath, vbDirectory) = vbNullString Then
            EnsureFolderRecursive parentPath
        End If
    End If

    MkDir folderPath
End Sub

' ==========================
' PATH RESOLUTION
' ==========================
Private Function GetExportBaseFolder() As String
    If IsUrlPath(ThisWorkbook.path) Or Len(ThisWorkbook.path) = 0 Then
        GetExportBaseFolder = Environ$("USERPROFILE") & "\OneDrive\Bureau"
    Else
        GetExportBaseFolder = ThisWorkbook.path
    End If
End Function

Private Function IsUrlPath(ByVal p As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(p))
    IsUrlPath = (Left$(s, 8) = "https://") Or (Left$(s, 7) = "http://")
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    SafeFileName = s
    For i = LBound(badChars) To UBound(badChars)
        SafeFileName = Replace(SafeFileName, badChars(i), "_")
    Next i
End Function


