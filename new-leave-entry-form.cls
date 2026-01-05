Option Compare Database

' Disable fields when the form loads
Private Sub Form_Load()
    On Error Resume Next

    Call DisableAllEmployeeFields

    ' Set default for EmployeeType
    Me.optEmployeeType = 1 ' 1 = PerEmployee, 2 = TemEmployee (adjust as needed)
End Sub

' Disable all employee fields initially
Private Sub DisableAllEmployeeFields()
    Me.cmbEmployeeName.Enabled = False
    Me.txtEmployeeID.Enabled = False
    Me.txtFolderID.Enabled = False
End Sub

' Handle Employee Input Method Selection
Private Sub optEmployeeInput_AfterUpdate()
    Call DisableAllEmployeeFields

    Select Case Nz(Me.optEmployeeInput.Value, 0)
        Case 1 ' Employee ID
            Me.txtEmployeeID.Enabled = True
        Case 2 ' Folder ID
            Me.txtFolderID.Enabled = True
        Case 3 ' Employee Name
            Me.cmbEmployeeName.Enabled = True
    End Select
End Sub

' Handle Employee Type Selection (Per / Tem)
Private Sub optEmployeeType_AfterUpdate()
    Dim strSQL As String

    If IsNull(Me.optEmployeeType) Then Exit Sub ' Prevent errors if no selection

    Select Case Me.optEmployeeType.Value
        Case 1
            strSQL = "SELECT EmployeeID, EmployeeFirstName & ' ' & EmployeeLastName, FolderID FROM PerEmployees ORDER BY EmployeeLastName;"
        Case 2
            strSQL = "SELECT EmployeeID, EmployeeFirstName & ' ' & EmployeeLastName, FolderID FROM TemEmployees ORDER BY EmployeeLastName;"
        Case Else
            Exit Sub
    End Select

    ' Update Employee ComboBox
    With Me.cmbEmployeeName
        .RowSource = strSQL
        .ColumnCount = 3
        .BoundColumn = 1
        .ColumnWidths = "0;3cm;0"
        .Requery
    End With
End Sub

' Auto-fill EmployeeID and FolderID when Name is selected
Private Sub cmbEmployeeName_AfterUpdate()
    If Not IsNull(Me.cmbEmployeeName) Then
        Me.txtEmployeeID.Value = Me.cmbEmployeeName.Column(0)  ' EmployeeID
        Me.txtFolderID.Value = Me.cmbEmployeeName.Column(2)  ' FolderID
        Call CalculateLeaveUsed  ' Recalculate leave when employee changes
    End If
End Sub

' Fill LeaveTypeID, AnnualLimit, and LeaveUnit when selecting a leave type
Private Sub cmbLeaveName_AfterUpdate()
    Dim rs As DAO.Recordset

    If IsNull(Me.cmbLeaveName) Then Exit Sub

    ' Fetch LeaveTypeID, AnnualLimit, and LeaveUnit from LeaveTypes table
    Set rs = CurrentDb.OpenRecordset("SELECT LeaveTypeID, AnnualLimit, LeaveUnit FROM LeaveTypes WHERE LeaveTypeID=" & Me.cmbLeaveName, dbOpenSnapshot)

    If Not rs.EOF Then
        Me.txtLeaveTypeID = rs!LeaveTypeID
        Me.txtAnnualLimit = rs!annualLimit
        Me.txtLeaveUnit = rs!LeaveUnit
    End If

    rs.Close
    Set rs = Nothing

    Call UpdateLeaveFields
    Call CalculateLeaveUsed
End Sub

' Lock/unlock HoursUsed and DaysUsed based on LeaveUnit
Private Sub UpdateLeaveFields()
    If Me.txtLeaveUnit = GetMessage("DayUnit") Then
        Me.txtHoursUsed.Value = Null
        Me.txtHoursUsed.Locked = True
        Me.txtDaysUsed.Locked = False
    ElseIf Me.txtLeaveUnit = GetMessage("HourUnit") Then
        Me.txtDaysUsed.Value = Null
        Me.txtDaysUsed.Locked = True
        Me.txtHoursUsed.Locked = False
End If
End Sub

' StartDate from frmCal
Private Sub btnSelectStartDate_Click()
    DoCmd.OpenForm "frmCal", , , , , acDialog
    If IsNumeric(STRDATE) Then
        Me.txtStartDate = CLng(STRDATE) ' Convert to number
        Call TriggerCalculation
    Else
        MsgBox "Invalid Date Selected", vbExclamation
    End If
End Sub

' EndDate from frmCal
Private Sub btnSelectEndDate_Click()
    DoCmd.OpenForm "frmCal", , , , , acDialog
    If IsNumeric(STRDATE) Then
        Me.txtEndDate = CLng(STRDATE) ' Convert to number
        Call TriggerCalculation
    Else
        MsgBox "Invalid Date Selected", vbExclamation
    End If
End Sub

' DateSaved from frmCal
Private Sub btnSelectDateSaved_Click()
    DoCmd.OpenForm "frmCal", , , , , acDialog
    If IsNumeric(STRDATE) Then
        Me.txtDateSaved = CLng(STRDATE)  ' Convert to a number
    Else
        MsgBox "Invalid Date Selected", vbExclamation
    End If
End Sub

' Auto-calculate DaysUsed when StartDate and EndDate are set
Private Sub txtEndDate_AfterUpdate()
    If Me.txtLeaveUnit = GetMessage("DayUnit") And Not IsNull(Me.txtStartDate) And Not IsNull(Me.txtEndDate) Then
        If Me.txtEndDate < Me.txtStartDate Then
            MsgBox GetMessage("InvalidDate"), vbExclamation, GetMessage("LeaveError")
            Me.txtEndDate = Null
            Exit Sub
        End If
        Me.txtDaysUsed = (Me.txtEndDate - Me.txtStartDate) + 1
        Me.txtHoursUsed = Null ' Ensure HoursUsed is empty for day-based leave
    End If
End Sub

' Auto-calculate DaysUsed when StartDate and EndDate are set
Private Sub TriggerCalculation()
    If IsNull(Me.txtLeaveUnit) Or IsNull(Me.txtStartDate) Or IsNull(Me.txtEndDate) Then Exit Sub

    If Me.txtLeaveUnit = GetMessage("DayUnit") Then
        If Me.txtEndDate < Me.txtStartDate Then
            MsgBox GetMessage("InvalidDate"), vbExclamation, GetMessage("LeaveError")
            Me.txtEndDate = Null
            Exit Sub
        End If
        Me.txtDaysUsed = (Me.txtEndDate - Me.txtStartDate) + 1
        Me.txtHoursUsed = Null
    End If

    Call CalculateLeaveUsed
End Sub

Function ConvertToDecimal(ByVal TimeStr As String) As Double
    Dim h As Integer, m As Integer
    On Error GoTo ErrHandler ' Handle invalid input gracefully
    
    If InStr(1, TimeStr, ":") > 0 Then
        h = Val(Left(TimeStr, InStr(1, TimeStr, ":") - 1)) ' Extract hours
        m = Val(Mid(TimeStr, InStr(1, TimeStr, ":") + 1)) ' Extract minutes
        ConvertToDecimal = h + (m / 60) ' Convert to decimal
    Else
        ConvertToDecimal = Val(TimeStr) ' If input is a pure number, assume hours
    End If
    Exit Function

ErrHandler:
    ConvertToDecimal = 0 ' Default to 0 on error
End Function

Function ConvertToHHMM(ByVal DecHours As Double) As String
    Dim h As Integer, m As Integer
    h = Int(DecHours) ' Get hours
    m = Round((DecHours - h) * 60, 0) ' Convert decimal part to minutes
    ConvertToHHMM = Format(h, "0") & ":" & Format(m, "00")
End Function

Private Sub txtHoursUsed_AfterUpdate()
    If Not IsNull(Me.txtHoursUsed) Then
        ' Convert input to decimal and store it as a Double in HiddenDecimalHours
        Me.HiddenDecimalHours = CDbl(ConvertToDecimal(Me.txtHoursUsed))
        
        ' Display the time in HH:MM format in txtHoursUsed
        Me.txtHoursUsed.Value = ConvertToHHMM(ConvertToDecimal(Me.txtHoursUsed))
        
        ' Update calculations
        Call CalculateLeaveUsed
    End If
End Sub

'Calculations
Private Sub CalculateLeaveUsed()
    Dim rs As DAO.Recordset
    Dim totalLeaveUsed As Double
    Dim leaveField As String
    Dim newLeaveUsed As Double

    If IsNull(Me.txtEmployeeID) Or IsNull(Me.txtLeaveTypeID) Then Exit Sub

    ' Determine field type
    If Me.txtLeaveUnit = GetMessage("DayUnit") Then
        leaveField = "DaysUsed"
    Else
        leaveField = "HoursUsed"
    End If

    ' Query database for total used leave
    Set rs = CurrentDb.OpenRecordset("SELECT SUM(" & leaveField & ") AS TotalUsed FROM LeaveRecords WHERE EmployeeID=" & Me.txtEmployeeID & " AND LeaveTypeID=" & Me.txtLeaveTypeID, dbOpenSnapshot)
    totalLeaveUsed = Nz(rs!TotalUsed, 0)
    rs.Close
    Set rs = Nothing

    ' Convert new leave entry (hh:mm) to decimal
    If Not IsNull(Me.txtDaysUsed) Then
        newLeaveUsed = Nz(Me.txtDaysUsed, 0)
    ElseIf Not IsNull(Me.txtHoursUsed) Then
        newLeaveUsed = ConvertToDecimal(Me.txtHoursUsed)
    Else
        newLeaveUsed = 0
    End If

    ' Sum total leave used
    totalLeaveUsed = totalLeaveUsed + newLeaveUsed

    ' Update fields
    If Me.txtLeaveUnit = GetMessage("HourUnit") Then
        Me.txtLeaveUsed = ConvertToHHMM(totalLeaveUsed) ' Show hh:mm format
        Me.txtLeaveRemained = ConvertToHHMM(Nz(Me.txtAnnualLimit, 0) - totalLeaveUsed) ' Show hh:mm format
    Else
        Me.txtLeaveUsed = totalLeaveUsed ' Show as plain number
        Me.txtLeaveRemained = Nz(Me.txtAnnualLimit, 0) - totalLeaveUsed ' Show as plain number
    End If

    ' Update leave status color
    If Nz(Me.txtAnnualLimit, 0) = 0 Then
        Me.txtLeaveStatus.BackColor = RGB(0, 0, 255) ' Blue for unlimited leave
    ElseIf Nz(Me.txtLeaveRemained, 0) >= 0.2 * Nz(Me.txtAnnualLimit, 0) Then
        Me.txtLeaveStatus.BackColor = RGB(0, 255, 0) ' Green (Normal)
    ElseIf Nz(Me.txtLeaveRemained, 0) > 0 Then
        Me.txtLeaveStatus.BackColor = RGB(255, 255, 0) ' Yellow (Near limit)
    Else
        Me.txtLeaveStatus.BackColor = RGB(255, 0, 0) ' Red (Exceeded limit)
    End If
End Sub

' Retrieve Persian Unit from Unit Table
Function GetMessage(MessageKey As String) As String
    Dim rs As DAO.Recordset
    Dim strSQL As String

    ' Fix SQL syntax
    strSQL = "SELECT MessageText FROM Unit WHERE MessageKey = '" & MessageKey & "'"

    ' Debugging: Check SQL statement
    Debug.Print strSQL

    ' Retrieve message
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
        GetMessage = rs!MessageText
    Else
        GetMessage = "[Message Not Found]"
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim newRecordID As Long
    Dim annualLimit As Double
    Dim leaveRemained As Double
    
    ' Set database connection
    Set db = CurrentDb()

    ' Ensure EmployeeType is selected
    If IsNull(Me.optEmployeeType) Or Me.optEmployeeType = 0 Then
        MsgBox "Please select an Employee Type (Per or Tem) before saving!", vbExclamation, "Missing Employee Type"
        Exit Sub
    End If

    ' Ensure StartDate is filled
    If IsNull(Me.txtStartDate) Or Me.txtStartDate = "" Then
        MsgBox "Start Date cannot be empty!", vbExclamation, "Missing Date"
        Exit Sub
    End If
    
    ' Ensure EndDate is filled
    If IsNull(Me.txtEndDate) Or Me.txtEndDate = "" Then
        MsgBox "End Date cannot be empty!", vbExclamation, "Missing Date"
        Exit Sub
    End If

    ' Ensure DateSaved is filled
    If IsNull(Me.txtDateSaved) Or Me.txtDateSaved = "" Then
        MsgBox "Date Saved cannot be empty!", vbExclamation, "Missing Date"
        Exit Sub
    End If

    ' Before generating the INSERT SQL
    If Me.txtLeaveUnit = GetMessage("DayUnit") Then
        Me.txtHoursUsed = Null
        Me.HiddenDecimalHours = Null
    ElseIf Me.txtLeaveUnit = GetMessage("HourUnit") Then
        Me.txtDaysUsed = Null
    End If
    
    'Prevent saving if Leave Used > Annual Limit
    annualLimit = Nz(Me.txtAnnualLimit, 0)
    If Me.txtLeaveUnit = GetMessage("HourUnit") Then
        leaveRemained = ConvertToDecimal(Nz(Me.txtLeaveRemained, "0:00")) ' Convert hh:mm to decimal
    Else
        leaveRemained = Val(Nz(Me.txtLeaveRemained, 0)) ' Keep as number for day-based leaves
    End If

    If annualLimit > 0 And leaveRemained < 0 Then
        MsgBox "Leave cannot be saved! Leave Used exceeds the Annual Limit.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Check for duplicate leave entry
    Set checkDuplicate = db.OpenRecordset("SELECT RecordID FROM LeaveRecords WHERE EmployeeID=" & Me.txtEmployeeID & _
        " AND LeaveTypeID=" & Me.cmbLeaveName & " AND StartDate=" & Me.txtStartDate, dbOpenSnapshot)

    If Not checkDuplicate.EOF Then
        MsgBox "Duplicate leave entry found! This leave has already been recorded.", vbExclamation, "Duplicate Entry"
        checkDuplicate.Close
        Set checkDuplicate = Nothing
        Exit Sub
    End If
    checkDuplicate.Close
    Set checkDuplicate = Nothing
          

    ' Construct SQL statement
    sql = "INSERT INTO LeaveRecords (EmployeeID, EmployeeType, FolderID, LeaveTypeID, StartDate, EndDate, HoursUsed, DaysUsed, DateSaved, Comment) " & _
      "VALUES (" & Nz(Me.txtEmployeeID, 0) & ", '" & Nz(Me.optEmployeeType.Value, "") & "', " & Nz(Me.txtFolderID, 0) & ", " & _
      Nz(Me.cmbLeaveName, 0) & ", " & Nz(Me.txtStartDate, 0) & ", " & Nz(Me.txtEndDate, 0) & ", " & _
      IIf(IsNull(Me.HiddenDecimalHours), "NULL", Format(Me.HiddenDecimalHours, "0.00")) & ", " & Nz(Me.txtDaysUsed, "NULL") & ", " & _
      Nz(Me.txtDateSaved, 0) & ", '" & Nz(Me.txtComment, "") & "');"
    
   ' Execute SQL statement
    db.Execute sql, dbFailOnError

    ' Retrieve the RecordID of the newly inserted record
    Set rs = db.OpenRecordset("SELECT @@IDENTITY AS NewID", dbOpenSnapshot)
    If Not rs.EOF Then
        newRecordID = rs!NewID
    End If
    rs.Close
    Set rs = Nothing

    ' Show success message with RecordID
    MsgBox "Leave entry saved successfully! Record ID: " & newRecordID, vbInformation, "Success"

    ' Refresh form and clear fields for new entry
    Me.Requery
    Me.optEmployeeType = Null
    Me.cmbEmployeeName = Null
    Me.txtEmployeeID = Null
    Me.txtFolderID = Null
    Me.cmbLeaveName = Null
    Me.txtLeaveTypeID = Null
    Me.txtLeaveUnit = Null
    Me.txtStartDate = Null
    Me.txtEndDate = Null
    Me.txtHoursUsed = Null
    Me.txtDaysUsed = Null
    Me.txtAnnualLimit = Null
    Me.txtLeaveUsed = Null
    Me.txtLeaveRemained = Null
    Me.txtLeaveStatus = Null
    Me.txtComment = Null

    ' Close database connection
    db.Close
    Set db = Nothing
    Exit Sub

Err_Handler:
    MsgBox "Error: " & err.Description, vbExclamation, "Save Error"
End Sub

Private Sub btnDelete_Click()
    Me.optEmployeeType = Null
    Me.cmbEmployeeName = Null
    Me.txtEmployeeID = Null
    Me.txtFolderID = Null
    Me.cmbLeaveName = Null
    Me.txtLeaveTypeID = Null
    Me.txtLeaveUnit = Null
    Me.txtStartDate = Null
    Me.txtEndDate = Null
    Me.txtHoursUsed = Null
    Me.txtDaysUsed = Null
    Me.txtDateSaved = Null
    Me.txtAnnualLimit = Null
    Me.txtLeaveUsed = Null
    Me.txtLeaveRemained = Null
    Me.txtLeaveStatus = Null
    Me.txtComment = Null
End Sub

