VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormulaireEntrees 
   Caption         =   "Demande de Remplacement :"
   ClientHeight    =   2325
   ClientLeft      =   315
   ClientTop       =   658
   ClientWidth     =   6377
   OleObjectBlob   =   "FormulaireEntrees.frx":0000
End
Attribute VB_Name = "FormulaireEntrees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' =============================================================================
' MODULE :          Code du UserForm
' DESCRIPTION :     Gère la boîte de dialogue de demande de remplacement.
'                   Aucune modification n'est nécessaire ici, mais inclus
'                   pour être complet.
' =============================================================================

' --- Module Level Variable ---
Private m_IsCancelled As Boolean

' --- Properties ---
Public Property Get IsCancelled() As Boolean
    IsCancelled = m_IsCancelled
End Property

Public Property Get SelectedEmployee() As String
    SelectedEmployee = Trim(Me.cboNomPrenom.value)
End Property

Public Property Get SelectedTeam() As String
    SelectedTeam = IIf(Me.optJour.value, "Jour", "Nuit")
End Property

Public Property Get PostCMValue() As String
    PostCMValue = IIf(Me.chkPostCM.value, "Post_CM", "")
End Property

Public Property Get ReplacementLines() As String
    ReplacementLines = Trim(Me.txtReplacementLines.value)
End Property

' --- Event Handlers ---
Private Sub UserForm_Initialize()
    On Error GoTo InitializationError
    m_IsCancelled = True
    Me.Caption = "Demande de Remplacement"
    Me.optJour.value = True
    Me.chkPostCM.value = False
    Me.chkMois.value = False
    LoadUniqueNames "Personnel", 2, 3, Me.cboNomPrenom
    chkMois_Click
    Me.cboNomPrenom.SetFocus
    Exit Sub
InitializationError:
    MsgBox "Erreur initialisation formulaire: " & Err.description, vbCritical
End Sub

Private Sub cmdOK_Click()
    On Error GoTo MainMacroError

    If Me.chkMois.value Then
        If Trim(Me.txtReplacementLines.value) = "" Then
            MsgBox "Veuillez entrer au moins un numéro de remplacement.", vbExclamation
            Me.txtReplacementLines.SetFocus
            Exit Sub
        End If
        m_IsCancelled = False
        Me.Hide
        GenerateNewWorkbookAndFillDates_Optimized_V4 _
            nomPrenom:="Demande remplacements Us 1D Jour et Nuit", _
            dayOrNight:="", _
            postCM:="/ MOIS", _
            ReplacementLines:=Me.ReplacementLines
    Else
        If Not ValidateInputs() Then Exit Sub
        m_IsCancelled = False
        Me.Hide
        GenerateNewWorkbookAndFillDates_Optimized_V4 _
            nomPrenom:=Me.SelectedEmployee, _
            dayOrNight:=Me.SelectedTeam, _
            postCM:=Me.PostCMValue, _
            ReplacementLines:=Me.ReplacementLines
    End If

    Unload Me
    Exit Sub
MainMacroError:
    ' C'est ce gestionnaire qui affichait le message simplifié.
    MsgBox "Erreur exécution macro principale: " & vbCrLf & Err.description, vbCritical
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    m_IsCancelled = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then m_IsCancelled = True
End Sub

Private Sub chkMois_Click()
    Dim isMonthlyMode As Boolean
    isMonthlyMode = Me.chkMois.value
    Me.cboNomPrenom.Enabled = Not isMonthlyMode
    Me.optJour.Enabled = Not isMonthlyMode
    Me.optNuit.Enabled = Not isMonthlyMode
    Me.chkPostCM.Enabled = Not isMonthlyMode
    If isMonthlyMode Then
        Me.cboNomPrenom.value = ""
        Me.chkPostCM.value = False
    End If
End Sub

' --- Helper Functions ---
Private Function ValidateInputs() As Boolean
    ValidateInputs = False
    If Me.cboNomPrenom.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un nom.", vbExclamation
        Me.cboNomPrenom.SetFocus
        Exit Function
    End If
    If Trim(Me.txtReplacementLines.value) = "" Then
        MsgBox "Veuillez entrer les numéros de ligne.", vbExclamation
        Me.txtReplacementLines.SetFocus
        Exit Function
    End If
    ValidateInputs = True
End Function

Private Sub LoadUniqueNames(ByVal sheetName As String, ByVal colFirstName As Long, ByVal colLastName As Long, ByRef targetCombo As MSForms.ComboBox)
    Dim ws As Worksheet, lastRow As Long, dataRange As Variant
    Dim namesDict As Object, i As Long, fullName As String, SortedKeys As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Feuille '" & sheetName & "' introuvable.", vbCritical
        Exit Sub
    End If
    lastRow = ws.Cells(ws.Rows.count, colFirstName).End(xlUp).row
    If lastRow < 2 Then Exit Sub
    dataRange = ws.Range(ws.Cells(2, colFirstName), ws.Cells(lastRow, colLastName)).value
    If Not IsArray(dataRange) Then Exit Sub
    Set namesDict = CreateObject("Scripting.Dictionary")
    namesDict.CompareMode = vbTextCompare
    For i = 1 To UBound(dataRange, 1)
        If Trim(dataRange(i, 1)) <> "" And Trim(dataRange(i, 2)) <> "" Then
            fullName = Trim(dataRange(i, 1)) & " " & Trim(dataRange(i, 2))
            If Not namesDict.Exists(fullName) Then namesDict.Add fullName, 1
        End If
    Next i
    If namesDict.count > 0 Then
        SortedKeys = SortDictionaryKeys(namesDict)
        targetCombo.list = SortedKeys
    End If
End Sub

Private Function SortDictionaryKeys(ByVal dict As Object) As Variant
    If dict.count = 0 Then
        SortDictionaryKeys = Array()
        Exit Function
    End If
    Dim keysArray() As String, i As Long, key As Variant
    ReDim keysArray(0 To dict.count - 1)
    i = 0
    For Each key In dict.keys
        keysArray(i) = CStr(key)
        i = i + 1
    Next key
    QuickSort keysArray, LBound(keysArray), UBound(keysArray)
    SortDictionaryKeys = keysArray
End Function

Private Sub QuickSort(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, pivot As String, temp As String
    If first >= last Then Exit Sub
    low = first
    high = last
    pivot = arr((first + last) \ 2)
    Do While low <= high
        Do While StrComp(arr(low), pivot, vbTextCompare) < 0
            low = low + 1
        Loop
        Do While StrComp(arr(high), pivot, vbTextCompare) > 0
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub

