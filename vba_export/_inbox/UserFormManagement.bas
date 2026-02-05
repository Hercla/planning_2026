' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "UserFormManagement"
Option Explicit

Sub ConfigureForms()
    Dim personalInfoFields As Variant
    personalInfoFields = Array( _
        Array("Matricule", "Matricule :"), _
        Array("Nom", "Nom :"), _
        Array("Prenom", "Prénom :"), _
        Array("Function", "Fonction :"), _
        Array("ContractType", "Type de contrat :"), _
        Array("Team", "Équipe :"), _
        Array("StartDate", "Date de début :"), _
        Array("EndDate", "Date de fin :"), _
        Array("BirthDate", "Date de naissance :"), _
        Array("Age", "Âge :"), _
        Array("EmailPrivate", "Email privé :"), _
        Array("EmailProfessional", "Email professionnel :") _
    )
    Call AddControlsToUserForm("ManagePersonalInfoForm", personalInfoFields, True)
    
    Dim leaveFields As Variant
    leaveFields = Array( _
        Array("Matricule", "Matricule :"), _
        Array("Nom", "Nom :"), _
        Array("Prenom", "Prénom :"), _
        Array("AncienneteJours", "Ancienneté Jours Théorique :"), _
        Array("AncienneteHeures", "Ancienneté Heures Théorique :"), _
        Array("CAJours", "CA Jours Théorique :"), _
        Array("CAHeures", "CA Heures Théorique :"), _
        Array("AFCJours", "AFC Jours Théorique :"), _
        Array("AFCHeures", "AFC Heures Théorique :"), _
        Array("ExtraLegauxJours", "Extra-Légaux Jours Théorique :"), _
        Array("ExtraLegauxHeures", "Extra-Légaux Heures Théorique :"), _
        Array("JoursFeries", "Jours fériés :"), _
        Array("CTRJours", "CTR Jours Théorique :"), _
        Array("CongeSocialJours", "Congé Social Jours Théorique :"), _
        Array("CongeSocialHeures", "Congé Social Heures Théorique :"), _
        Array("CEPJours", "CEP Jours Théorique :"), _
        Array("CEPHeures", "CEP Heures Théorique :"), _
        Array("VacanceJeuneJours", "Vacance Jeune Jours Théorique :"), _
        Array("VacanceJeuneHeures", "Vacance Jeune Heures Théorique :") _
    )
    Call AddControlsToUserForm("ManageLeavesForm", leaveFields, False)
    
    Dim baseFields As Variant
    baseFields = Array( _
        Array("Matricule", "Matricule :"), _
        Array("Nom", "Nom :"), _
        Array("Prenom", "Prénom :") _
    )
    
    ' Generate position fields
    Dim positionFields As Variant
    positionFields = MergeArrays(baseFields, GenerateMonthFields("Position", "Position"))
    Call AddControlsToUserForm("ManagePositionsForm", positionFields, False, 343, 441)
    
    ' Generate work time fields
    Dim workTimeFields As Variant
    workTimeFields = MergeArrays(baseFields, GenerateMonthFields("WorkTime", "%"))
    Call AddControlsToUserForm("ManageWorkTimeForm", workTimeFields, False)
End Sub

Function GenerateMonthFields(suffix As String, captionSuffix As String) As Variant
    Dim months As Variant
    months = Array("Janv", "Fevr", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    Dim fields() As Variant
    ReDim fields(0 To UBound(months))
    Dim i As Integer
    For i = 0 To UBound(months)
        fields(i) = Array(months(i) & suffix, months(i) & " " & captionSuffix & " :")
    Next i
    GenerateMonthFields = fields
End Function

Function MergeArrays(arr1 As Variant, arr2 As Variant) As Variant
    Dim arr() As Variant
    Dim i As Long, n As Long, m As Long
    n = UBound(arr1) - LBound(arr1) + 1
    m = UBound(arr2) - LBound(arr2) + 1
    ReDim arr(0 To n + m - 1)
    For i = 0 To n - 1
        arr(i) = arr1(i)
    Next i
    For i = 0 To m - 1
        arr(n + i) = arr2(i)
    Next i
    MergeArrays = arr
End Function

Sub AddControlsToUserForm(frmName As String, fields As Variant, hasDatePicker As Boolean, Optional height As Integer = 0, Optional width As Integer = 0)
    Dim frm As Object
    Dim ctrl As Object
    Dim topPosition As Integer
    Dim leftPosition As Integer
    Dim i As Integer
    Dim ctrlName As String
    
    ' Set the UserForm
    Set frm = ThisWorkbook.VBProject.VBComponents(frmName).Designer
    
    ' Clear existing controls
    For i = frm.Controls.Count - 1 To 0 Step -1
        frm.Controls.Remove frm.Controls(i).Name
    Next i
    
    ' Start positions
    topPosition = 10
    leftPosition = 10
    
    ' Add fields
    For i = LBound(fields) To UBound(fields)
        ' Add Label
        ctrlName = "lbl" & fields(i)(0)
        Set ctrl = frm.Controls.Add("Forms.Label.1", ctrlName)
        With ctrl
            .Caption = fields(i)(1)
            .Top = topPosition
            .Left = leftPosition
            .width = 100
        End With
        
        ' Determine control type
        Select Case fields(i)(0)
            Case "Nom", "Prenom", "Function", "ContractType", "Team"
                ctrlName = "cmb" & fields(i)(0)
                Set ctrl = frm.Controls.Add("Forms.ComboBox.1", ctrlName)
            Case Else
                ctrlName = "txt" & fields(i)(0)
                Set ctrl = frm.Controls.Add("Forms.TextBox.1", ctrlName)
        End Select
        With ctrl
            .Top = topPosition
            .Left = leftPosition + 110
            .width = 100
        End With
        
        ' Increment top position for the next row
        topPosition = topPosition + 30
        
        ' Reset position for next column after a certain number of fields
        If (i + 1) Mod 10 = 0 Then
            topPosition = 10
            leftPosition = leftPosition + 250
        End If
    Next i
    
    ' Add DatePicker Buttons if necessary
    If hasDatePicker Then
        AddDatePickerButton frm, "txtStartDate", "btnStartDate"
        AddDatePickerButton frm, "txtEndDate", "btnEndDate"
        ' Add Frame for DatePicker
        AddDatePickerFrame frm
    End If
    
    ' Add Submit button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnSubmit")
    With ctrl
        .Caption = "Submit"
        .Top = frm.InsideHeight - 50
        .Left = frm.InsideWidth - 180
        .width = 80
    End With
    
    ' Add Close button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnClose")
    With ctrl
        .Caption = "Fermer"
        .Top = frm.InsideHeight - 50
        .Left = frm.InsideWidth - 90
        .width = 80
    End With
End Sub

Function ControlExists(frm As Object, ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = frm.Controls(ctrlName)
    ControlExists = Not ctrl Is Nothing
    On Error GoTo 0
End Function

Sub AddDatePickerButton(frm As Object, txtBoxName As String, btnName As String)
    If Not ControlExists(frm, btnName) Then
        Dim ctrl As Object
        Set ctrl = frm.Controls.Add("Forms.CommandButton.1", btnName)
        With ctrl
            .Caption = "..."
            .Top = frm.Controls(txtBoxName).Top
            .Left = frm.Controls(txtBoxName).Left + frm.Controls(txtBoxName).width + 10
            .width = 20
            .height = frm.Controls(txtBoxName).height
        End With
    End If
End Sub

Sub AddDatePickerFrame(frm As Object)
    Dim ctrl As Object
    Dim frame As Object
    Dim clsDay As clsCalendarDay
    
    ' Initialiser la collection des jours si elle n'est pas déjà définie
    If dayLabels Is Nothing Then
        Set dayLabels = New Collection
    End If
    
    ' Ajout du cadre pour le sélecteur de date si nécessaire
    If Not ControlExists(frm, "frameDatePicker") Then
        Set frame = frm.Controls.Add("Forms.Frame.1", "frameDatePicker")
        With frame
            .Caption = ""
            .Visible = False
            .width = 250
            .height = 250
            .Top = 10
            .Left = 650
        End With
        
        ' Ajout des labels pour les jours de la semaine
        Dim daysOfWeek As Variant
        daysOfWeek = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
        Dim i As Integer
        For i = LBound(daysOfWeek) To UBound(daysOfWeek)
            Set ctrl = frame.Controls.Add("Forms.Label.1", "lblDay" & i)
            With ctrl
                .Caption = daysOfWeek(i)
                .Top = 20
                .Left = 20 + (i * 30)
                .width = 30
            End With
        Next i
        
        ' Ajout des labels pour les jours du mois
        Dim j As Integer, lblName As String
        For i = 1 To 6
            For j = 1 To 7
                lblName = "day" & ((i - 1) * 7 + j)
                Set ctrl = frame.Controls.Add("Forms.Label.1", lblName)
                With ctrl
                    .Caption = ""
                    .Top = 40 + ((i - 1) * 30)
                    .Left = 20 + ((j - 1) * 30)
                    .width = 30
                    .height = 20
                    .BackColor = &HFFFFFF
                    .BorderStyle = fmBorderStyleSingle
                    .TextAlign = fmTextAlignCenter
                End With
                ' Assigner chaque label à une instance de clsCalendarDay
                Set clsDay = New clsCalendarDay
                Set clsDay.lblDay = ctrl
                Set clsDay.ParentForm = frm
                ' Ajouter l'instance dans la collection globale
                dayLabels.Add clsDay
            Next j
        Next i
        
        ' Ajout des boutons de navigation
        Set ctrl = frame.Controls.Add("Forms.CommandButton.1", "btnLastMonth")
        With ctrl
            .Caption = "<"
            .Top = 10
            .Left = 20
            .width = 30
        End With
        
        Set ctrl = frame.Controls.Add("Forms.CommandButton.1", "btnNextMonth")
        With ctrl
            .Caption = ">"
            .Top = 10
            .Left = 180
            .width = 30
        End With
    End If
End Sub


' DatePicker functions
Sub BuildCalendar(frm As Object, Optional iYear As Integer, Optional iMonth As Integer)
    Dim startOfMonth As Date
    Dim trackingDate As Date
    Dim iStartofMonthDay As Integer
    Dim cDay As Control
    
    If iYear = 0 Or iMonth = 0 Then
        iYear = Year(Now())
        iMonth = Month(Now())
    End If
    
    With frm
        .Controls("frameDatePicker").Controls("lblMonth").Caption = monthName(iMonth, True)
        .Controls("frameDatePicker").Controls("lblYear").Caption = iYear
        
        startOfMonth = DateSerial(iYear, iMonth, 1)
        iStartofMonthDay = Weekday(startOfMonth, vbSunday)
        trackingDate = DateAdd("d", -iStartofMonthDay + 1, startOfMonth)
        
        For i = 1 To 42
            Set cDay = .Controls("frameDatePicker").Controls("day" & i)
            cDay.Caption = Day(trackingDate)
            cDay.Tag = trackingDate
            
            If Month(trackingDate) <> iMonth Then
                cDay.ForeColor = &H808080
            Else
                cDay.ForeColor = &H0
            End If
            
            trackingDate = DateAdd("d", 1, trackingDate)
        Next i
    End With
End Sub

Sub DayClick(frm As Object, selectedDate As Date, dateControl As Control)
    dateControl.text = selectedDate
    frm.Controls("frameDatePicker").Visible = False
End Sub

Sub ToggleDatePicker(frm As Object, oControl As Control)
    If frm.Controls("frameDatePicker").Visible Then
        frm.Controls("frameDatePicker").Visible = False
    Else
        If IsDate(oControl.text) Then
            BuildCalendar frm, Year(oControl.text), Month(oControl.text)
        Else
            BuildCalendar frm
        End If
        ' Store the control to update later
        frm.Tag = oControl.Name
        frm.Controls("frameDatePicker").Top = oControl.Top + oControl.height + 5
        frm.Controls("frameDatePicker").Left = oControl.Left
        frm.Controls("frameDatePicker").Visible = True
    End If
End Sub

Sub FillEmployeeComboBoxes()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Personnel")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Clear and fill the ComboBoxes for each UserForm
    Dim userForms As Variant
    userForms = Array("ManageLeavesForm", "ManagePositionsForm", "ManageWorkTimeForm", "ManagePersonalInfoForm")
    
    Dim frm As Object
    Dim j As Integer
    For j = LBound(userForms) To UBound(userForms)
        Set frm = ThisWorkbook.VBProject.VBComponents(userForms(j)).Designer
        With frm
            If ControlExists(frm, "cmbNom") Then
                .Controls("cmbNom").Clear
                .Controls("cmbPrenom").Clear
                For i = 2 To lastRow
                    .Controls("cmbNom").AddItem ws.Cells(i, 2).value ' Nom
                    .Controls("cmbPrenom").AddItem ws.Cells(i, 3).value ' Prénom
                Next i
            End If
        End With
    Next j
End Sub

' Class Module: clsCalendarDay
Option Explicit

Public WithEvents lblDay As MSForms.Label
Public ParentForm As Object

Private Sub lblDay_Click()
    Dim frm As Object
    Set frm = ParentForm
    Dim selectedDate As Date
    selectedDate = lblDay.Tag
    Dim dateControl As Control
    Set dateControl = frm.Controls(frm.Tag)
    DayClick frm, selectedDate, dateControl
End Sub

' Usage in UserForm Module
Private Sub btnStartDate_Click()
    ToggleDatePicker Me, Me.Controls("txtStartDate")
End Sub

Private Sub btnEndDate_Click()
    ToggleDatePicker Me, Me.Controls("txtEndDate")
End Sub

Private Sub cmbNom_Change()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Personnel")
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 2).value = Me.Controls("cmbNom").value Then
            Me.Controls("cmbPrenom").value = ws.Cells(i, 3).value
            Exit For
        End If
    Next i
End Sub

Private Sub cmbPrenom_Change()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Personnel")
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 3).value = Me.Controls("cmbPrenom").value Then
            Me.Controls("cmbNom").value = ws.Cells(i, 2).value
            Exit For
        End If
    Next i
End Sub

Private Sub UserForm_Initialize()
    FillEmployeeComboBoxes
End Sub


