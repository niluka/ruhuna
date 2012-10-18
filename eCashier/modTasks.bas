Attribute VB_Name = "modTasks"
Option Explicit
    Dim temSql As String

Public Sub SaveCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                SaveSetting App.EXEName, MyForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i)
            Next
        End If
    Next
    SaveSetting App.EXEName, MyForm.Name, "Top", MyForm.Top
    SaveSetting App.EXEName, MyForm.Name, "Left", MyForm.Left
End Sub

Public Sub GetCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                MyCtrl.ColWidth(i) = GetSetting(App.EXEName, MyForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i))
                MyCtrl.AllowUserResizing = flexResizeColumns
            Next
        End If
    Next
    MyForm.Top = GetSetting(App.EXEName, MyForm.Name, "Top", MyForm.Top)
    MyForm.Left = GetSetting(App.EXEName, MyForm.Name, "Left", MyForm.Left)
End Sub

Public Sub GridToExcel(ExportGrid As MSFlexGrid, Optional Topic As String, Optional Subtopic As String)
    On Error Resume Next
    
    If ExportGrid.Rows <= 1 Then
        MsgBox "Noting to Export"
        Exit Sub
    End If
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myWorkSheet1 As Excel.Worksheet
    Dim temRow As Integer
    Dim temCol As Integer
    
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set myWorkSheet1 = AppExcel.Worksheets(1)
    
    myWorkSheet1.Cells(1, 1) = Topic
    myWorkSheet1.Cells(2, 1) = Subtopic
    
    For temRow = 0 To ExportGrid.Rows - 1
        For temCol = 0 To ExportGrid.Cols - 1
            myWorkSheet1.Cells(temRow + 3, temCol + 1) = ExportGrid.TextMatrix(temRow, temCol)
            If ExportGrid.ColWidth(temCol) < 5 Then
                myWorkSheet1.Columns(, temCol + 1).Hidden = True
            End If
        Next
    Next temRow
    
    myworkbook.SaveAs (App.Path & "\" & Topic & ".xls")
    myworkbook.Save
    myworkbook.Close
    
    ShellExecute 0&, "open", App.Path & "\" & Topic & ".xls", "", "", vbMaximizedFocus
End Sub

Public Function ProfessionalFeePaidBHT(BillID As Long) As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "SELECT dbo.tblProfessionalCharges.PaidDate FROM dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForBHTID = dbo.tblIncomeBill.BHTID WHERE     (dbo.tblIncomeBill.IncomeBillID = " & BillID & " AND dbo.tblProfessionalCharges.paid = 1)"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ProfessionalFeePaidBHT = True
        Else
            ProfessionalFeePaidBHT = False
        End If
    End With
End Function

Public Function ProfessionalFeePaidGSB(BillID As Long) As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "SELECT dbo.tblProfessionalCharges.PaidDate FROM dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForBHTID = dbo.tblIncomeBill.BHTID WHERE     (dbo.tblIncomeBill.IncomeBillID = " & BillID & " AND dbo.tblProfessionalCharges.paid = 1)"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ProfessionalFeePaidGSB = True
        Else
            ProfessionalFeePaidGSB = False
        End If
    End With
End Function


Public Function ProfessionalFeePaidOPD(BillID As Long) As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "SELECT  dbo.tblIncomeBill.IncomeBillID, dbo.tblProfessionalCharges.PaidDate FROM dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForOPDBillID = dbo.tblIncomeBill.IncomeBillID WHERE     (dbo.tblIncomeBill.IncomeBillID = " & BillID & " AND dbo.tblProfessionalCharges.paid = 1)"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ProfessionalFeePaidOPD = True
        Else
            ProfessionalFeePaidOPD = False
        End If
    End With
End Function

Public Function ProfessionalFeePaidLab(BillID As Long) As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "SELECT  dbo.tblIncomeBill.IncomeBillID, dbo.tblProfessionalCharges.PaidDate FROM dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForLabBillID = dbo.tblIncomeBill.IncomeBillID WHERE     (dbo.tblIncomeBill.IncomeBillID = " & BillID & " AND dbo.tblProfessionalCharges.Paid = 1)"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ProfessionalFeePaidLab = True
        Else
            ProfessionalFeePaidLab = False
        End If
    End With
End Function

Public Function ProfessionalFeePaidMedicalTest(BillID As Long) As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "SELECT  dbo.tblIncomeBill.IncomeBillID, dbo.tblProfessionalCharges.PaidDate FROM dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForMedicalTestBillID = dbo.tblIncomeBill.IncomeBillID WHERE     (dbo.tblIncomeBill.IncomeBillID = " & BillID & " AND dbo.tblProfessionalCharges.Paid = 1)"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ProfessionalFeePaidMedicalTest = True
        Else
            ProfessionalFeePaidMedicalTest = False
        End If
    End With
End Function

Public Function ProfessionalFeePaidHST(BillID As Long) As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "SELECT  dbo.tblIncomeBill.IncomeBillID, dbo.tblProfessionalCharges.PaidDate FROM dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForHSTBillID = dbo.tblIncomeBill.IncomeBillID WHERE     (dbo.tblIncomeBill.IncomeBillID = " & BillID & " AND dbo.tblProfessionalCharges.Paid = 1)"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ProfessionalFeePaidHST = True
        Else
            ProfessionalFeePaidHST = False
        End If
    End With
End Function

Public Function ProfessionalFeePaidR(BillID As Long) As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "SELECT  dbo.tblIncomeBill.IncomeBillID, dbo.tblProfessionalCharges.PaidDate FROM dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForRBillID = dbo.tblIncomeBill.IncomeBillID WHERE     (dbo.tblIncomeBill.IncomeBillID = " & BillID & " AND dbo.tblProfessionalCharges.Paid = 1)"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ProfessionalFeePaidR = True
        Else
            ProfessionalFeePaidR = False
        End If
    End With
End Function


Public Function FullStaffName(StaffID As Long) As String
    Dim rsTem As New ADODB.Recordset
    Dim temTitleID As Long
    
    With rsTem
'        temSql = "SELECT     dbo.tblTitle.Title + SPACE(1) + dbo.tblStaff.Name AS FullStaffName " & _
'                    "FROM         dbo.tblStaff LEFT OUTER JOIN " & _
'                      "dbo.tblTitle ON dbo.tblStaff.TitleID = dbo.tblTitle.TitleID " & _
'                    "WHERE     (dbo.tblStaff.StaffID = " & StaffID & ")"
        
        
        If .State = 1 Then .Close
        temSql = "SELECT     dbo.tblStaff.Name AS FullStaffName, tblStaff.TitleID " & _
                    "FROM         dbo.tblStaff LEFT OUTER JOIN " & _
                      "dbo.tblTitle ON dbo.tblStaff.TitleID = dbo.tblTitle.TitleID " & _
                    "WHERE     (dbo.tblStaff.StaffID = " & StaffID & ")"
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!FullStaffName) = False Then
                FullStaffName = !FullStaffName
                temTitleID = !TitleID
            Else
                FullStaffName = "Staff"
            End If
        Else
            FullStaffName = "Staff"
        End If
        
        If .State = 1 Then .Close
        temSql = "SELECT     dbo.tblTitle.Title  " & _
                    "FROM tblTitle " & _
                    "WHERE tblTitle.TitleID = " & temTitleID & ""
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Title) = False Then
                FullStaffName = !Title & " " & FullStaffName
            End If
        End If
        
    End With
End Function

Public Function FullPatientName(PatientID As Long) As String
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "SELECT     tblPatientMainDetails.* " & _
                    "FROM         tblPatientMainDetails " & _
                    "WHERE     (PatientID = " & PatientID & ")"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!FirstName) = False Then
                FullPatientName = !FirstName
            Else
                FullPatientName = "Customer"
            End If
        Else
            FullPatientName = "Customer"
        End If
    End With
End Function

