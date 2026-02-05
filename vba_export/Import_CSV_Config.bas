Attribute VB_Name = "Import_CSV_Config"
Option Explicit

' ========================================================================================
' SCRIPT D'IMPORT AUTOMATIQUE DES CONFIGURATIONS
' A executer une fois pour mettre a jour Feuil_Config et Config_Codes
' ========================================================================================

Sub MettreAJourConfig()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Chemins des fichiers CSV (A adapter si besoin, mais devraient etre au meme endroit)
    Dim pathConfig As String
    Dim pathCodes As String
    
    ' On suppose que les fichiers CSV sont dans le meme dossier que le fichier Excel
    ' OU dans le dossier specifie c:\Users\hercl\planning_2026\
    pathConfig = "c:\Users\hercl\planning_2026\Feuil_Config_CORRIGE.csv"
    pathCodes = "c:\Users\hercl\planning_2026\Config_Codes_COMPLET.csv"
    
    If Dir(pathConfig) = "" Then
        MsgBox "Fichier introuvable : " & pathConfig, vbCritical
        Exit Sub
    End If
    If Dir(pathCodes) = "" Then
        MsgBox "Fichier introuvable : " & pathCodes, vbCritical
        Exit Sub
    End If
    
    ' 1. Mettre a jour Feuil_Config
    ImporterCSV wb, "Feuil_Config", pathConfig
    
    ' 2. Mettre a jour Config_Codes
    ImporterCSV wb, "Config_Codes", pathCodes
    
    MsgBox "Mise a jour terminee avec succes !" & vbCrLf & _
           "- Feuil_Config : OK" & vbCrLf & _
           "- Config_Codes : OK", vbInformation
End Sub

Private Sub ImporterCSV(wb As Workbook, sheetName As String, csvPath As String)
    Dim ws As Worksheet
    Dim numFile As Integer
    Dim lineStr As String
    Dim lineArr() As String
    Dim rowNum As Long
    Dim colNum As Integer
    Dim content As String
    
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Feuille '" & sheetName & "' introuvable !", vbExclamation
        Exit Sub
    End If
    
    ' Vider la feuille
    ws.Cells.Clear
    
    ' Lire le fichier CSV
    numFile = FreeFile
    Open csvPath For Input As #numFile
    
    rowNum = 1
    Do Until EOF(numFile)
        Line Input #numFile, lineStr
        
        ' Gestion simple du point-virgule
        lineArr = Split(lineStr, ";")
        
        For colNum = 0 To UBound(lineArr)
            content = Trim(lineArr(colNum))
            ' Enlever les guillemets si presents
            If Left(content, 1) = """" And Right(content, 1) = """" Then
                content = Mid(content, 2, Len(content) - 2)
            End If
            
            ws.Cells(rowNum, colNum + 1).value = content
        Next colNum
        
        rowNum = rowNum + 1
    Loop
    
    Close #numFile
    
    ' Mise en forme rapide
    ws.Columns("A:Z").AutoFit
End Sub
